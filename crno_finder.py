# -*- coding: utf-8 -*-
"""
법인등록번호 찾기 (금융위원회 기업기본정보 API / 공공데이터포털)

입력 : 엑셀 (기업명, 대표자명 열)
출력 : <입력파일명>_등록번호.xlsx  (법인등록번호, 판정상태, 후보, 주소 등 추가)

사용법
  pip install pandas openpyxl
  python crno_finder.py --test-key                       # API 키/응답 구조 점검
  python crno_finder.py --input "목록.xlsx" --limit 5    # 5건만 시험
  python crno_finder.py --input "목록.xlsx"              # 전체 실행
  python crno_finder.py --input "목록.xlsx" --resume     # 중단 후 이어서

API 키 지정 (우선순위): --api-key  >  환경변수 DATA_GO_KR_KEY  >  이 파일과 같은 폴더의 key.txt
"""
import argparse
import json
import os
import re
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

import pandas as pd

API_URL = "https://apis.data.go.kr/1160100/service/GetCorpBasicInfoService_V2/getCorpOutline_V2"
COMPANY_ALIASES = ["기업명", "회사명", "법인명", "상호", "업체명", "기업", "회사", "회사명"]
REP_ALIASES = ["대표자명", "대표자", "대표이사", "대표", "대표자성명", "성명", "이름"]

ST_OK = "확정"
ST_MULTI = "동명다수(대표자도 동일)"
ST_REP_MISS = "상호일치-대표자불일치"
ST_REP_NONE = "상호일치-대표자정보없음"
ST_NAME_MISS = "상호불일치"
ST_NOT_FOUND = "API에 없음"
ST_ERR = "오류"

DELAY = 0.25  # API 호출 간 대기(초)


# ============================ 문자열 처리 ============================
def norm_company(s):
    s = str(s or "")
    s = re.sub(r"\(주\)|㈜|주식회사|\(유\)|유한회사|유한책임회사|\(합\)|합자회사|합명회사", "", s)
    return re.sub(r"[\s\W_]+", "", s).lower()


def core_name(s):
    """검색용 핵심 상호 (법인형태 표기 제거, 공백 유지)"""
    s = re.sub(r"\(주\)|㈜|주식회사|\(유\)|유한회사|유한책임회사", "", str(s or ""))
    return re.sub(r"\s+", " ", s).strip()


def rep_tokens(s):
    """대표자 문자열 -> 비교용 이름 집합 (공백 제거 전체 + 구분자로 나눈 토큰)"""
    s = str(s or "").strip()
    if not s:
        return set()
    out = {re.sub(r"\s+", "", s)}
    s2 = re.sub(r"외\s*\d+\s*(인|명)", " ", s)
    for t in re.split(r"[,/·、&\s]+|및", s2):
        t = t.strip()
        if len(t) >= 2:
            out.add(t)
    return out


def rep_match(excel_rep, api_rep):
    a, b = rep_tokens(excel_rep), rep_tokens(api_rep)
    return bool(a & b)


def hyphen(crno):
    c = re.sub(r"\D", "", str(crno or ""))
    return f"{c[:6]}-{c[6:]}" if len(c) == 13 else c


# ============================ API ============================
class ApiError(Exception):
    pass


def load_key(cli_key):
    key = cli_key or os.environ.get("DATA_GO_KR_KEY", "")
    tried = []
    if not key:
        for base in (Path(__file__).resolve().parent, Path.cwd()):
            for name in ("key.txt", "key.txt.txt"):
                kp = base / name
                tried.append(str(kp))
                if kp.exists():
                    for enc in ("utf-8-sig", "utf-16", "cp949"):
                        try:
                            key = kp.read_text(encoding=enc).strip().strip('"').strip("'")
                            break
                        except Exception:
                            continue
                    if key:
                        break
            if key:
                break
    if not key:
        sys.exit("API 키를 찾지 못했습니다. 확인한 위치:\n  " + "\n  ".join(tried or ["(없음)"]) +
                 "\n--api-key, 환경변수 DATA_GO_KR_KEY, 또는 key.txt 중 하나로 지정하세요.")
    return urllib.parse.unquote(key.strip())  # 인코딩된 키/디코딩된 키 모두 허용


def fetch_page(key, corp_nm, page=1, rows=100):
    q = urllib.parse.urlencode({
        "serviceKey": key, "pageNo": page, "numOfRows": rows,
        "resultType": "json", "corpNm": corp_nm,
    })
    last = None
    for attempt in range(3):
        try:
            with urllib.request.urlopen(f"{API_URL}?{q}", timeout=20) as r:
                raw = r.read().decode("utf-8", errors="replace")
            try:
                data = json.loads(raw)
            except json.JSONDecodeError:
                raise ApiError("JSON이 아닌 응답(키 오류/미승인 가능): " + re.sub(r"\s+", " ", raw)[:200])
            resp = data.get("response", data)
            header = resp.get("header", {})
            code = str(header.get("resultCode", "00"))
            if code != "00":
                raise ApiError(f"API 오류 {code}: {header.get('resultMsg', '')}")
            body = resp.get("body", {})
            items = body.get("items") or {}
            items = items.get("item", []) if isinstance(items, dict) else []
            if isinstance(items, dict):
                items = [items]
            return items, int(body.get("totalCount", len(items)) or 0)
        except ApiError:
            raise
        except (urllib.error.URLError, TimeoutError, ConnectionError) as e:
            last = e
            time.sleep(1.5 * (attempt + 1))
    raise ApiError(f"네트워크 오류: {last}")


_cache = {}


def fetch_all(key, corp_nm, max_pages=5):
    if corp_nm in _cache:
        return _cache[corp_nm]
    items, page = [], 1
    while page <= max_pages:
        got, total = fetch_page(key, corp_nm, page)
        items += got
        time.sleep(DELAY)
        if len(items) >= total or not got:
            break
        page += 1
    _cache[corp_nm] = items
    return items


# ============================ 매칭 ============================
def find_crno(key, company, rep):
    """반환 dict: status, crno, candidates(str), 그리고 확정 시 부가정보"""
    target = norm_company(company)
    queries = []
    for q in (company.strip(), core_name(company), "주식회사 " + core_name(company)):
        if q and q not in queries:
            queries.append(q)

    pool = {}
    for q in queries:
        for it in fetch_all(key, q):
            pool[(it.get("crno"), it.get("corpNm"))] = it
        if any(norm_company(i.get("corpNm", "")) == target for i in pool.values()):
            break

    items = list(pool.values())
    if not items:
        return {"status": ST_NOT_FOUND, "crno": "", "cand": ""}

    exact = [i for i in items if norm_company(i.get("corpNm", "")) == target]

    def cand_text(lst):
        return " | ".join(
            f"{i.get('corpNm','')}({i.get('enpRprFnm','') or '대표자?'}) {i.get('crno','')}"
            for i in lst[:5]
        )

    if not exact:
        return {"status": ST_NAME_MISS, "crno": "", "cand": cand_text(items)}

    if not str(rep or "").strip():
        crnos = {i["crno"] for i in exact if i.get("crno")}
        return {"status": "대표자미입력-" + ("후보1건" if len(crnos) == 1 else "후보다수"),
                "crno": "", "cand": cand_text(exact)}

    hit = [i for i in exact if rep_match(rep, i.get("enpRprFnm", ""))]
    crnos = sorted({i["crno"] for i in hit if i.get("crno")})
    if len(crnos) == 1:
        it = next(i for i in hit if i.get("crno") == crnos[0])
        return {"status": ST_OK, "crno": crnos[0], "cand": "", "item": it}
    if len(crnos) > 1:
        return {"status": ST_MULTI, "crno": "", "cand": cand_text(hit)}
    if all(not str(i.get("enpRprFnm", "")).strip() for i in exact):
        return {"status": ST_REP_NONE, "crno": "", "cand": cand_text(exact)}
    return {"status": ST_REP_MISS, "crno": "", "cand": cand_text(exact)}


# ============================ 엑셀 ============================
def pick_column(df, aliases, label):
    cols = {str(c).strip(): c for c in df.columns}
    for a in aliases:
        if a in cols:
            return cols[a]
    for c in cols:  # 부분 일치
        if any(a in c for a in aliases):
            return cols[c]
    sys.exit(f"'{label}' 열을 찾지 못했습니다. 현재 열: {list(df.columns)}")


def save_excel(df, path):
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="결과")
        ws = w.sheets["결과"]
        ws.freeze_panes = "A2"
        from openpyxl.styles import Font, PatternFill
        for c in ws[1]:
            c.font = Font(bold=True)
            c.fill = PatternFill("solid", fgColor="DDEBF7")
        for i, col in enumerate(df.columns, start=1):
            width = max(len(str(col)) * 2, *(min(len(str(v)), 50) for v in df[col].head(200))) + 2
            ws.column_dimensions[ws.cell(1, i).column_letter].width = min(width, 60)
        status_col = list(df.columns).index("판정상태") + 1
        for r in range(2, len(df) + 2):
            cell = ws.cell(r, status_col)
            color = "C6EFCE" if cell.value == ST_OK else ("FFEB9C" if cell.value != ST_ERR else "FFC7CE")
            cell.fill = PatternFill("solid", fgColor=color)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--input")
    ap.add_argument("--output")
    ap.add_argument("--api-key", default="")
    ap.add_argument("--sheet", default=0, help="시트 이름 또는 번호(0부터)")
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--resume", action="store_true", help="기존 결과 파일에서 '확정' 행은 건너뛰고 이어서")
    ap.add_argument("--test-key", action="store_true", help="API 키 점검")
    args = ap.parse_args()

    key = load_key(args.api_key)

    if args.test_key:
        try:
            items, total = fetch_page(key, "삼성전자", 1, 3)
        except ApiError as e:
            sys.exit(f"[실패] {e}")
        print(f"[성공] 총 {total}건. 첫 결과:")
        print(json.dumps(items[0] if items else {}, ensure_ascii=False, indent=2))
        return

    if not args.input:
        sys.exit("--input 엑셀 파일을 지정하세요.")
    sheet = int(args.sheet) if str(args.sheet).isdigit() else args.sheet
    df = pd.read_excel(args.input, sheet_name=sheet, dtype=str).fillna("")
    c_company = pick_column(df, COMPANY_ALIASES, "기업명")
    c_rep = pick_column(df, REP_ALIASES, "대표자명")
    print(f"사용 열: 기업명='{c_company}', 대표자명='{c_rep}', 총 {len(df)}행")

    out_cols = ["법인등록번호", "법인등록번호(하이픈)", "판정상태", "후보", "API상호", "API대표자", "주소", "사업자등록번호", "설립일"]
    out_path = Path(args.output) if args.output else Path(args.input).with_name(Path(args.input).stem + "_등록번호.xlsx")

    if args.resume and out_path.exists():
        prev = pd.read_excel(out_path, dtype=str).fillna("")
        if len(prev) == len(df):
            df = prev
            print(f"이어하기: {out_path}")
    for c in out_cols:
        if c not in df.columns:
            df[c] = ""

    idxs = list(df.index[: args.limit]) if args.limit else list(df.index)
    n_ok = 0
    try:
        for n, i in enumerate(idxs, 1):
            company, rep = df.at[i, c_company].strip(), df.at[i, c_rep].strip()
            if not company or df.at[i, "판정상태"] == ST_OK:
                n_ok += df.at[i, "판정상태"] == ST_OK
                continue
            try:
                r = find_crno(key, company, rep)
            except ApiError as e:
                if "키" in str(e) or "API 오류" in str(e):
                    save_excel(df, out_path)
                    sys.exit(f"\n치명적 오류로 중단(지금까지 결과 저장됨): {e}")
                r = {"status": ST_ERR, "crno": "", "cand": str(e)[:150]}
            df.at[i, "판정상태"] = r["status"]
            df.at[i, "법인등록번호"] = r["crno"]
            df.at[i, "법인등록번호(하이픈)"] = hyphen(r["crno"])
            df.at[i, "후보"] = r.get("cand", "")
            it = r.get("item")
            if it:
                df.at[i, "API상호"] = it.get("corpNm", "")
                df.at[i, "API대표자"] = it.get("enpRprFnm", "")
                df.at[i, "주소"] = it.get("enpBsadr", "")
                df.at[i, "사업자등록번호"] = it.get("bzno", "")
                df.at[i, "설립일"] = it.get("enpEstbDt", "")
            n_ok += r["status"] == ST_OK
            print(f"[{n}/{len(idxs)}] {company} / {rep} -> {r['status']} {r['crno']}")
            if n % 20 == 0:
                save_excel(df, out_path)
    except KeyboardInterrupt:
        print("\n사용자 중단 - 지금까지 결과를 저장합니다.")

    save_excel(df, out_path)
    st = df.loc[idxs, "판정상태"].value_counts()
    print("\n===== 요약 =====")
    for k, v in st.items():
        print(f"  {k or '(미처리)'}: {v}")
    print(f"확정 {n_ok}/{len(idxs)}건 -> {out_path}")


if __name__ == "__main__":
    main()