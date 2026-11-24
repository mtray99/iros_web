# -*- coding: utf-8 -*-
"""
법인등기 열람 자동신청 웹 프로그램
  엑셀(회사명, 대표자) 업로드 → 법인등록번호 조회(금융위 API) → 인터넷등기소 결제대상 목록에 담기
  ※ 로그인과 결제는 반드시 사람이 직접 합니다. (결제 버튼은 절대 누르지 않음)

실행:  python app.py   → 브라우저에서 http://127.0.0.1:5000
"""
import contextlib
import re
import threading
import time
import traceback
import uuid
import webbrowser
from pathlib import Path

import pandas as pd
from flask import Flask, jsonify, render_template, request, send_file
from playwright.sync_api import sync_playwright

import crno_finder as cf
import iros_apply_v7 as iros

BASE = Path(__file__).resolve().parent
OUT = BASE / "output"
OUT.mkdir(exist_ok=True)
app = Flask(__name__)

CRNO_COLS = ["법인등록번호", "법인등록번호(하이픈)", "판정상태", "후보", "API상호", "API대표자", "주소", "사업자등록번호", "설립일"]
APPLY_COLS = ["신청상태", "신청비고", "선택된행", "장바구니건수"]


# ============================ 작업 상태 ============================
class Timer:
    """누적 시간 측정 (일시정지 가능: 로그인/결제 대기 시간은 제외)"""

    def __init__(self):
        self.acc, self.t = 0.0, None

    def start(self):
        if self.t is None:
            self.t = time.time()

    def stop(self):
        if self.t is not None:
            self.acc += time.time() - self.t
            self.t = None

    def elapsed(self):
        return self.acc + (time.time() - self.t if self.t else 0.0)


class Job:
    def __init__(self):
        self.id = uuid.uuid4().hex[:8]
        self.phase = "idle"   # idle lookup wait_login apply wait_batch wait_pay done stopped error
        self.msg = ""
        self.log = []
        self.df = None
        self.c_company = self.c_rep = None
        self.out_path = None
        self.rev = 0
        self.prog = {"label": "", "done": 0, "total": 0}
        self.timers = {"lookup": Timer(), "apply": Timer()}
        self.running = False
        self.stop = threading.Event()        # 중지: 처리 루프 중단
        self.login_evt = threading.Event()   # 사람이 로그인 완료
        self.pay_evt = threading.Event()     # 중간 결제 후 계속
        self.finish_evt = threading.Event()  # 마지막 결제 후 종료

    def touch(self):
        self.rev += 1


JOB = Job()


class LogWriter:
    """print() 출력을 작업 로그로 가져오는 stdout 대용품"""

    def __init__(self, job):
        self.job, self.buf = job, ""

    def write(self, s):
        self.buf += s
        while "\n" in self.buf:
            line, self.buf = self.buf.split("\n", 1)
            if line.strip():
                self.job.log.append(line.rstrip())
        return len(s)

    def flush(self):
        pass


def wait_evt(job, evt, honor_stop=True):
    """이벤트가 올 때까지 대기. 중지 요청이 오면 False"""
    while not evt.wait(0.4):
        if honor_stop and job.stop.is_set():
            return False
    return True


def save(job):
    cf.save_excel(job.df, job.out_path)
    job.touch()


# ============================ 1단계: 법인등록번호 조회 ============================
def run_lookup(job, o):
    df = job.df
    idxs = list(df.index[: o["limit"]]) if o["limit"] else list(df.index)
    todo = [i for i in idxs if df.at[i, job.c_company].strip() and df.at[i, "판정상태"] != cf.ST_OK]
    job.prog = {"label": "법인등록번호 조회", "done": 0, "total": len(todo)}
    if not todo:
        print("조회할 행이 없습니다(이미 '확정').")
        return
    try:
        key = cf.load_key(o["api_key"])
    except SystemExit as e:
        raise RuntimeError(str(e))
    for n, i in enumerate(todo, 1):
        if job.stop.is_set():
            print("사용자 중지")
            break
        company = df.at[i, job.c_company].strip()
        rep = df.at[i, job.c_rep].strip() if job.c_rep else ""
        try:
            r = cf.find_crno(key, company, rep)
        except cf.ApiError as e:
            if "키" in str(e) or "API 오류" in str(e):
                save(job)
                raise RuntimeError(f"API 오류로 중단(지금까지 결과 저장됨): {e}")
            r = {"status": cf.ST_ERR, "crno": "", "cand": str(e)[:150]}
        df.at[i, "판정상태"], df.at[i, "법인등록번호"] = r["status"], r["crno"]
        df.at[i, "법인등록번호(하이픈)"] = cf.hyphen(r["crno"])
        df.at[i, "후보"] = r.get("cand", "")
        it = r.get("item")
        if it:
            df.at[i, "API상호"], df.at[i, "API대표자"] = it.get("corpNm", ""), it.get("enpRprFnm", "")
            df.at[i, "주소"], df.at[i, "사업자등록번호"] = it.get("enpBsadr", ""), it.get("bzno", "")
            df.at[i, "설립일"] = it.get("enpEstbDt", "")
        print(f"[조회 {n}/{len(todo)}] {company} / {rep} -> {r['status']} {r['crno']}")
        job.prog["done"] = n
        job.touch()
        if n % 20 == 0:
            save(job)
    save(job)


# ============================ 2단계: 인터넷등기소 신청 ============================
def run_apply(job, o, page, ctx):
    df = job.df

    def eligible(i):
        if not re.sub(r"\D", "", df.at[i, "법인등록번호"]):
            return False
        if df.at[i, "판정상태"] != cf.ST_OK and not o["include_unconfirmed"]:
            return False
        return df.at[i, "신청상태"] != iros.ST_DONE

    targets = [i for i in df.index if eligible(i)]
    if o["limit"]:
        targets = targets[: o["limit"]]
    print(f"신청 대상 {len(targets)}건 / 전체 {len(df)}행")
    job.prog = {"label": "인터넷등기소 신청", "done": 0, "total": len(targets)}
    opt = {"dry_run": o["dry_run"], "pick_first": o["pick_first"], "scope": o["scope"],
           "items": o["items"], "list_items": False}

    seen, last_count, added = set(), -1, 0
    for n, i in enumerate(targets, 1):
        if job.stop.is_set():
            print("사용자 중지")
            break
        company = df.at[i, job.c_company].strip()
        crno = re.sub(r"\D", "", df.at[i, "법인등록번호"])
        rep = df.at[i, job.c_rep].strip() if job.c_rep else ""
        job.prog["done"] = n - 1
        job.msg = f"{company} 처리 중"
        job.touch()
        if crno in seen:
            df.at[i, "신청상태"], df.at[i, "신청비고"] = "중복(건너뜀)", "같은 등록번호가 앞 행에 있음"
            continue
        seen.add(crno)
        print(f"[신청 {n}/{len(targets)}] {company} ({crno})")
        st = note = picked = ""
        for attempt in (1, 2):
            try:
                iros.close_extra_pages(ctx, page)
                st, note, picked, cnt = iros.apply_one(page, company, crno, opt, last_count, rep)
                if st == iros.ST_DONE and cnt >= 0:
                    last_count = cnt
                    df.at[i, "장바구니건수"] = str(cnt)
                break
            except Exception as e:
                st, note, picked = iros.ST_ERR, str(e)[:200], ""
                iros.dismiss(page)
                if attempt == 1:
                    print(f"    (오류, 재시도) {note[:80]}")
                    page.wait_for_timeout(1000)
        df.at[i, "신청상태"], df.at[i, "신청비고"], df.at[i, "선택된행"] = st, note, picked
        print(f"    -> {st} {note}")
        save(job)

        if st == iros.ST_DONE:
            added += 1
            if o["batch"] and added % o["batch"] == 0:
                cnt = iros.open_pay_list(page)
                job.phase = "wait_batch"
                job.msg = f"{added}건 담음 (결제대상 목록 {cnt if cnt is not None else '확인불가'}건). 결제 후 [계속]을 누르세요."
                job.touch()
                job.pay_evt.clear()
                job.timers["apply"].stop()          # 결제 대기 시간은 소요시간에서 제외
                ok = wait_evt(job, job.pay_evt)
                job.timers["apply"].start()
                if not ok:
                    break
                job.phase, last_count = "apply", -1
    job.prog["done"] = job.prog["total"] if not job.stop.is_set() else job.prog["done"]
    return added


# ============================ 전체 파이프라인 ============================
def pipeline(job, o):
    w = LogWriter(job)
    try:
        with contextlib.redirect_stdout(w):
            use_browser = not o["crno_only"]
            pw = browser = None
            try:
                if use_browser:
                    pw = sync_playwright().start()
                    browser = pw.chromium.launch(headless=False, slow_mo=50, args=["--window-size=1400,900"])
                    ctx = browser.new_context(viewport={"width": 1400, "height": 900}, locale="ko-KR")
                    page = ctx.new_page()
                    page.set_default_timeout(15000)

                    def on_dialog(d):
                        print(f"    (알림창/{d.type}) {d.message[:80]}")
                        d.dismiss() if (d.type == "confirm" and "결제" in d.message) else d.accept()

                    page.on("dialog", on_dialog)
                    page.goto("https://www.iros.go.kr/index.jsp", wait_until="domcontentloaded", timeout=30000)
                    page.wait_for_timeout(2500)
                    iros.close_extra_pages(ctx, page)
                    print("인터넷등기소 창을 열었습니다. 조회가 진행되는 동안 그 창에서 로그인을 미리 해두세요.")

                job.phase, job.msg = "lookup", "법인등록번호 조회 중"
                job.touch()
                job.timers["lookup"].start()
                try:
                    run_lookup(job, o)
                finally:
                    job.timers["lookup"].stop()
                if job.stop.is_set():
                    job.phase, job.msg = "stopped", "중지됨"
                    return
                if not use_browser:
                    job.phase, job.msg = "done", "법인등록번호 조회 완료"
                    return

                n_ok = int((job.df["판정상태"] == cf.ST_OK).sum())
                print(f"조회 완료: 확정 {n_ok}/{len(job.df)}건")
                job.phase = "wait_login"
                job.msg = "인터넷등기소 창에서 로그인(아이디/자동입력방지문자)을 마치고 [로그인 완료]를 누르세요."
                job.touch()
                if not wait_evt(job, job.login_evt):
                    job.phase, job.msg = "stopped", "중지됨"
                    return
                iros.close_extra_pages(ctx, page)

                job.phase, job.msg = "apply", "신청 진행 중"
                job.touch()
                job.timers["apply"].start()
                try:
                    added = run_apply(job, o, page, ctx)
                finally:
                    job.timers["apply"].stop()

                if not o["dry_run"] and added:
                    cnt = iros.open_pay_list(page)
                    job.phase = "wait_pay"
                    job.msg = (f"결제대상 목록 {cnt}건. " if cnt and cnt > 0 else "") + \
                              "열린 창에서 직접 결제하세요(법인은 페이지당 최대 10건). 끝나면 [종료]를 누르세요."
                    job.touch()
                    wait_evt(job, job.finish_evt, honor_stop=False)
                job.phase = "stopped" if job.stop.is_set() else "done"
                job.msg = "중지됨" if job.stop.is_set() else "완료"
            finally:
                if browser:
                    try:
                        browser.close()
                    except Exception:
                        pass
                if pw:
                    try:
                        pw.stop()
                    except Exception:
                        pass
                if job.df is not None:
                    save(job)
    except Exception as e:
        job.log.append(traceback.format_exc())
        job.phase, job.msg = "error", str(e)
    finally:
        job.running = False
        job.touch()


# ============================ 라우트 ============================
def parse_items(s):
    return [int(x) for x in re.split(r"[,\s]+", s or "") if x.strip().isdigit()]


@app.get("/")
def index():
    return render_template("index.html")


@app.post("/api/start")
def start():
    global JOB
    if JOB.running:
        return jsonify(error="이미 실행 중인 작업이 있습니다."), 409
    f = request.files.get("file")
    if not f or not f.filename.lower().endswith((".xlsx", ".xls")):
        return jsonify(error="엑셀 파일(.xlsx)을 첨부하세요."), 400
    form = request.form
    o = {
        "api_key": form.get("api_key", "").strip(),
        "limit": 0,
        "scope": "partial",                       # '일부' 클릭 + 항목 12번 행이 기본
        "items": parse_items(form.get("items")),
        "pick_first": form.get("pick_first") == "1",
        "dry_run": False,
        "crno_only": form.get("crno_only") == "1",
        "include_unconfirmed": False,
        "batch": 0,
    }
    o["items"] = o["items"] or iros.SCOPE_DEFAULT_ITEMS["partial"]

    job = Job()
    d = OUT / job.id
    d.mkdir(parents=True)
    src = d / Path(f.filename).name
    f.save(src)
    try:
        df = pd.read_excel(src, dtype=str).fillna("")
        job.c_company = cf.pick_column(df, cf.COMPANY_ALIASES, "회사명")
        job.c_rep = cf.pick_column(df, cf.REP_ALIASES, "대표자명")
    except SystemExit as e:
        return jsonify(error=str(e)), 400
    except Exception as e:
        return jsonify(error=f"엑셀을 읽지 못했습니다: {e}"), 400
    for c in CRNO_COLS + APPLY_COLS:
        if c not in df.columns:
            df[c] = ""
    job.df = df
    job.out_path = d / (src.stem.replace("_등록번호", "") + "_결과.xlsx")
    job.log.append(f"파일: {src.name} / 회사명 열='{job.c_company}', 대표자 열='{job.c_rep}', {len(df)}행")
    job.running = True
    JOB = job
    threading.Thread(target=pipeline, args=(job, o), daemon=True).start()
    return jsonify(ok=True)


@app.get("/api/status")
def status():
    job, since, rev = JOB, int(request.args.get("since", 0)), int(request.args.get("rev", -1))
    res = {"phase": job.phase, "msg": job.msg, "running": job.running, "prog": job.prog,
           "rev": job.rev, "logn": len(job.log), "log": job.log[since:], "job": job.id,
           "times": {k: round(t.elapsed()) for k, t in job.timers.items()},
           "timing": [k for k, t in job.timers.items() if t.t is not None],
           "can_download": bool(job.out_path and job.out_path.exists())}
    if job.df is not None and rev != job.rev:
        df = job.df
        res["rows"] = [[df.at[i, job.c_company], df.at[i, job.c_rep] if job.c_rep else "",
                        df.at[i, "법인등록번호(하이픈)"], df.at[i, "판정상태"], df.at[i, "신청상태"],
                        df.at[i, "신청비고"] or df.at[i, "후보"]] for i in df.index]
        res["counts"] = {"ok": int((df["판정상태"] == cf.ST_OK).sum()),
                         "done": int((df["신청상태"] == iros.ST_DONE).sum()), "total": len(df)}
    return jsonify(res)


@app.post("/api/action")
def action():
    a = (request.get_json(silent=True) or {}).get("action")
    if a == "login_done":
        JOB.login_evt.set()
    elif a == "resume":
        JOB.pay_evt.set()
    elif a == "finish":
        JOB.finish_evt.set()
    elif a == "stop":
        JOB.stop.set()
        JOB.login_evt.set()
        JOB.pay_evt.set()
    else:
        return jsonify(error="알 수 없는 동작"), 400
    return jsonify(ok=True)


@app.get("/api/download")
def download():
    if not (JOB.out_path and JOB.out_path.exists()):
        return jsonify(error="결과 파일이 아직 없습니다."), 404
    return send_file(JOB.out_path, as_attachment=True, download_name=JOB.out_path.name)


if __name__ == "__main__":
    url = "http://127.0.0.1:5000"
    threading.Timer(1.0, lambda: webbrowser.open(url)).start()
    print(f"서버 시작: {url}  (종료: Ctrl+C)")
    app.run(host="127.0.0.1", port=5000, threaded=True)
