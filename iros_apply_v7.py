# -*- coding: utf-8 -*-
"""
인터넷등기소 법인등기 열람 신청 매크로 v7 (법인등록번호 검색 → 결제대상 목록에 담기)

입력 : crno_finder.py 결과 엑셀 (법인등록번호, 판정상태='확정' 인 행만 처리)
동작 : [사람] 로그인 → [매크로] 등록번호 검색 → 결과 행 선택 → 범위/항목 선택 → 결제대상 목록에 담기 (회사마다 반복)
       [사람] 마지막에 결제대상 목록에서 직접 결제   ※ 결제 버튼은 절대 누르지 않습니다

사용법
  python iros_apply_v7.py --input "목록_등록번호.xlsx" --dry-run --limit 3      # 신청 없이 행 선택까지 확인
  python iros_apply_v7.py --input "목록_등록번호.xlsx" --limit 2 --debug        # 2건 실제 진행 + 화면 저장
  python iros_apply_v7.py --input "목록_등록번호.xlsx" --resume                 # 전체 / 이어서
  옵션: --scope partial|malso   --items 12   --pick-first   --batch 10
"""
import argparse
import re
import sys
import time
import traceback
from pathlib import Path

import pandas as pd
from playwright.sync_api import sync_playwright

VERSION = "v7.2 (임원 전체선택 단계 + 대표자 확인 추가, 2026-11-19)"
P = "mf_wfm_potal_main_wfm_content_"

IDS = {
    "menu": "mf_wfm_potal_main_wf_header_gen_depth1_0_gen_depth2_1_gen_depth3_0_btn_top_menu3a",
    "tab_crno": P + "tac_crg_srch_tab_tab_drokno_tabHTML",
    "input_crno": P + "sbx_drokno___input",
    "btn_search": P + "btn_dorkno_search",
    "btn_next": P + "btn_next",
    "btn_pay": P + "btn_pay",
    "btn_add": P + "btn_new_add",
    "btn_pay_list": P + "btn_pay_list",
    "kind_malso": P + "rad_crg_kind_input_1",           # '말소사항 포함' 라벨(for=)
    "item_grid": P + "grd_item_sel_obj_list",
    "item_chk": "G_" + P + "grd_item_sel_obj_list___checkbox_dynamic_checkbox_{i}_0_{i}",
}

# 범위 선택 방식별 기본 항목 행: 사용자가 녹화한 방식('일부' + 12번 행)이 기본
SCOPE_DEFAULT_ITEMS = {"partial": [12], "malso": [14, 15]}
SCOPE_TEXT = "일부"

ST_DONE = "결제대기(신청완료)"
ST_DRY = "행확인(dry-run)"
ST_NOROW = "검색결과없음"
ST_MULTI = "행다수-수동확인"
ST_ERR = "오류"

DEBUG = False
DEBUG_DIR = Path("output/debug")
COMPANY_ALIASES = ["기업명", "회사명", "법인명", "상호", "업체명", "기업", "회사"]
REP_ALIASES = ["대표자명", "대표자", "대표이사", "대표", "대표자성명"]
NOSELECT = [0]            # '선택한 사항이 없습니다' 안내창 발생 횟수
LAST_ALERTS = []          # 방금 처리한 알림창 문구들 (검색결과 없음 판단용)


# ============================ 유틸 ============================
def pick_company_column(df):
    cols = [str(c).strip() for c in df.columns]
    for a in COMPANY_ALIASES:
        if a in cols:
            return df.columns[cols.index(a)]
    for i, c in enumerate(cols):
        if any(a in c for a in COMPANY_ALIASES):
            return df.columns[i]
    sys.exit(f"기업명 열을 찾지 못했습니다. 현재 열: {list(df.columns)}")


def pick_optional_column(df, aliases):
    cols = [str(c).strip() for c in df.columns]
    for a in aliases:
        if a in cols:
            return df.columns[cols.index(a)]
    return None


def name_in_text(name, text):
    """한글 이름이 다른 한글 글자에 붙지 않은 경우만 일치로 인정"""
    name = re.sub(r"\s+", "", str(name or ""))
    return bool(name) and re.search(rf"(?<![가-힣]){re.escape(name)}(?![가-힣])", text) is not None


def norm_company(s):
    s = re.sub(r"\(주\)|㈜|주식회사|\(유\)|유한회사|유한책임회사", "", str(s or ""))
    return re.sub(r"[\s\W_]+", "", s).lower()


def dump(page, tag):
    """--debug 일 때 화면/텍스트 저장"""
    if not DEBUG:
        return
    DEBUG_DIR.mkdir(parents=True, exist_ok=True)
    try:
        page.screenshot(path=str(DEBUG_DIR / f"{tag}.png"), full_page=True)
        (DEBUG_DIR / f"{tag}.txt").write_text(page.locator("body").inner_text(), encoding="utf-8")
    except Exception:
        pass


def close_extra_pages(ctx, keep):
    for pg in list(ctx.pages):
        if pg != keep:
            try:
                pg.close()
            except Exception:
                pass


def js_click(page, el_id):
    return page.evaluate("id => { const e = document.getElementById(id); if (!e) return false; e.click(); return true; }", el_id)


def has_id(page, el_id):
    return page.evaluate("id => !!document.getElementById(id)", el_id)


# ============================ 팝업 처리 ============================
# 화면에 떠 있는 '확인창(message_popup)'을 처리한다.
#  - 안내문에 '결제'가 있고 [취소] 버튼이 있으면 → [취소] (= 결제하지 않고 신청 건을 추가). [확인]은 절대 누르지 않음
#  - 버튼이 1개뿐인 알림창 → 그 버튼으로 닫음
#  - 그 밖의 확인창(버튼 2개 이상) → 건드리지 않고 내용만 보고
DISMISS_JS = r"""() => {
  const vis = e => { const r = e.getBoundingClientRect(); const s = getComputedStyle(e);
    return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none'; };
  const label = b => (b.innerText || b.value || '').trim();
  const btns = [...document.querySelectorAll('a[id*="message_popup"], button[id*="message_popup"], input[id*="message_popup"]')].filter(vis);
  if (!btns.length) return { popup: false, out: [] };
  const textOf = b => { let e = b; for (let i = 0; i < 9; i++) { e = e.parentElement; if (!e) break;
    const t = (e.innerText || '').replace(/\s+/g, ' ').trim(); if (t.length >= 12) return t; } return ''; };
  const groups = {};
  btns.forEach(b => { const k = b.id.replace(/_wframe.*$/, ''); (groups[k] = groups[k] || []).push(b); });
  const out = [];
  for (const k of Object.keys(groups)) {
    const bs = groups[k]; const t = textOf(bs[0]);
    const cancel = bs.find(b => /btn_cancel/.test(b.id) && label(b) === '취소');
    if (/선택한 사항이 없/.test(t) && cancel) { cancel.click(); out.push({ action: 'cancel_noselect', text: t, names: bs.map(label) }); continue; }
    if (/결제/.test(t) && cancel) { cancel.click(); out.push({ action: 'cancel_pay', text: t, names: bs.map(label) }); continue; }
    if (bs.length === 1) { bs[0].click(); out.push({ action: 'alert_ok', text: t, names: [label(bs[0])] }); continue; }
    out.push({ action: 'unknown', text: t, names: bs.map(label) });
  }
  return { popup: true, out };
}"""

HIDE_MODAL_JS = r"""() => { document.querySelectorAll('#_modal,.w2modal_popup').forEach(m => { m.style.display = 'none'; m.style.pointerEvents = 'none'; }); }"""
_unknown_seen = set()


def dismiss(page):
    handled = []
    for fr in page.frames:
        try:
            res = fr.evaluate(DISMISS_JS)
        except Exception:
            continue
        if not res.get("popup"):
            try:
                fr.evaluate(HIDE_MODAL_JS)   # 팝업 없이 남은 반투명 오버레이 제거
            except Exception:
                pass
            continue
        for o in res["out"]:
            handled.append(o)
            if o["action"] == "cancel_pay":
                print("    (결제할 증명서 안내창 -> [취소]=추가 선택)")
                LAST_ALERTS.append(o["text"])
            elif o["action"] == "cancel_noselect":
                NOSELECT[0] += 1
                print("    ('선택한 사항이 없습니다' 안내창 -> [취소]로 되돌아감)")
            elif o["action"] == "alert_ok":
                print(f"    (알림창 닫음 [{o['names'][0]}]) {o['text'][:100]}")
                LAST_ALERTS.append(o["text"])
            elif (o["text"], tuple(o["names"])) not in _unknown_seen:
                _unknown_seen.add((o["text"], tuple(o["names"])))
                print(f"    (처리하지 않은 확인창) {o['text'][:140]} / 버튼: {o['names']}")
                dump(page, "POPUP_unknown")
    return handled


def wait_for(page, fn, timeout=15000, step=300):
    end = time.time() + timeout / 1000
    while time.time() < end:
        dismiss(page)
        try:
            r = fn()
        except Exception:
            r = None
        if r:
            return r
        page.wait_for_timeout(step)
    return None


# ============================ 화면 동작 ============================
ROWS_JS = r"""() => [...document.querySelectorAll('input[type="radio"][id*="grd_srch_rslt_list"], input[type="checkbox"][id*="grd_srch_rslt_list"]')]
  .map((r, i) => { const row = r.closest('tr') || r.parentElement; const raw = row ? row.innerText : '';
    return { idx: i, text: raw.replace(/\s+/g, ' ').trim(), cells: raw.split('\t').map(x => x.trim()) }; })"""

PICK_JS = r"""(i) => { const rs = [...document.querySelectorAll('input[type="radio"][id*="grd_srch_rslt_list"], input[type="checkbox"][id*="grd_srch_rslt_list"]')];
  if (!rs[i]) return false; rs[i].click(); return !!rs[i].checked; }"""

SET_JS = r"""(a) => { const inp = document.getElementById(a.id); if (!inp) return null;
  inp.focus(); inp.value = ''; inp.dispatchEvent(new Event('input', { bubbles: true }));
  inp.value = a.num; inp.dispatchEvent(new Event('input', { bubbles: true }));
  inp.dispatchEvent(new Event('keyup', { bubbles: true })); inp.dispatchEvent(new Event('change', { bubbles: true }));
  return inp.value; }"""

STATE_JS = r"""(a) => {
  const g = id => document.getElementById(id);
  const vis = e => { const r = e.getBoundingClientRect(); return r.width > 0 && r.height > 0; };
  const own = e => [...e.childNodes].filter(n => n.nodeType === 3).map(n => n.textContent).join('').trim();
  const scope = [...document.querySelectorAll('label,span,a,div,td,li')].some(e => vis(e) && own(e) === a.scope);
  return { hasPay: !!g(a.pay), hasAdd: !!g(a.add), hasScope: scope, hasKind: !!document.querySelector('label[for="' + a.kind + '"]'),
           hasSelAll: [...document.querySelectorAll('input[type=checkbox]')].some(i => (i.getAttribute('aria-label') || '').trim() === '전체선택') || [...document.querySelectorAll('label,span,th,div,td')].some(e => vis(e) && own(e) === '전체선택'),
           hasItems: !!g(a.item0), hasNext: !!g(a.next) };
}"""

CLICK_TEXT_JS = r"""(t) => {
  const own = e => [...e.childNodes].filter(n => n.nodeType === 3).map(n => n.textContent).join('').trim();
  const vis = e => { const r = e.getBoundingClientRect(); return r.width > 0 && r.height > 0; };
  const el = [...document.querySelectorAll('label,span,a,div,td,li')].find(e => vis(e) && own(e) === t);
  if (!el) return false; el.click(); return true; }"""

CHECK_ITEMS_JS = r"""(a) => a.rows.map(i => { const c = document.getElementById(a.tpl.split('{i}').join(i));
  if (!c) return [i, 'missing']; if (!c.checked) c.click(); return [i, !!c.checked]; })"""

ITEM_ROWS_JS = r"""(grid) => { const rows = {}; document.querySelectorAll('[id^="' + grid + '_cell_"]').forEach(e => {
  const m = e.id.match(/_cell_(\d+)_(\d+)$/); if (!m) return; const r = +m[1], c = +m[2];
  (rows[r] = rows[r] || {})[c] = (e.innerText || '').replace(/\s+/g, ' ').trim(); });
  return Object.keys(rows).map(Number).sort((a, b) => a - b).map(r => [r, Object.keys(rows[r]).map(Number).sort((a, b) => a - b).map(c => rows[r][c]).filter(Boolean).join(' | ')]); }"""

LABEL_CLICK_JS = r"""(a) => { const l = document.querySelector('#' + a.grid + '_cell_' + a.i + '_0 > .w2checkbox_label'); if (!l) return false; l.click(); return true; }"""


HEADER_LEFT_JS = r"""() => {   // '직책' 헤더 바로 왼쪽 칸의 체크박스(전체선택)를 클릭
  const own = e => [...e.childNodes].filter(n => n.nodeType === 3).map(n => n.textContent).join('').trim();
  const th = [...document.querySelectorAll('th,td,div,span')].find(e => own(e) === '직책');
  if (!th) return false;
  let cell = th.closest('th,td') || th, prev = cell.previousElementSibling;
  while (prev) { const c = prev.querySelector('input[type=checkbox], .w2checkbox_label, label'); if (c) { c.click(); return true; } prev = prev.previousElementSibling; }
  return false; }"""

ANY_CHECKED_JS = r"""() => [...document.querySelectorAll('input[type=checkbox]:checked')].length"""


def select_all_officers(page):
    """임원 목록 화면의 '전체선택'(헤더 체크박스, 직책 왼쪽)을 체크한다. 성공하면 방식 문자열, 실패하면 None"""
    cb = page.get_by_role("checkbox", name="전체선택")

    def header_checked():
        try:
            return cb.count() > 0 and cb.first.is_checked()
        except Exception:
            return False

    if header_checked():
        return "이미 체크됨"
    try:                                           # 1) 녹화와 같은 방식: 체크박스를 직접 체크
        if cb.count() > 0:
            cb.first.check(force=True, timeout=3000)
            if header_checked():
                return "체크박스"
    except Exception:
        pass
    for loc in (page.get_by_role("columnheader", name="전체선택"), page.get_by_text("전체선택", exact=True)):
        try:                                       # 2) 헤더/문구 클릭
            if loc.count() > 0:
                loc.first.click(timeout=2500, force=True)
                if header_checked() or (cb.count() == 0 and page.evaluate(ANY_CHECKED_JS) > 0):
                    return "헤더클릭"
        except Exception:
            pass
    try:                                           # 3) '직책' 왼쪽 칸의 버튼
        if page.evaluate(HEADER_LEFT_JS):
            page.wait_for_timeout(300)
            if header_checked() or page.evaluate(ANY_CHECKED_JS) > 0:
                return "직책왼쪽"
    except Exception:
        pass
    return None


def go_to_search(page):
    """상단 메뉴 → 등록번호검색 탭 (메뉴를 다시 누르면 검색 화면이 초기화됨)"""
    dismiss(page)
    if not js_click(page, IDS["menu"]):
        raise RuntimeError("상단 메뉴 버튼을 찾지 못함 (로그인 상태/화면 확인)")
    page.wait_for_timeout(2000)
    dismiss(page)
    page.wait_for_timeout(800)
    dismiss(page)
    if not wait_for(page, lambda: has_id(page, IDS["tab_crno"])):
        raise RuntimeError("등록번호검색 탭을 찾지 못함")
    js_click(page, IDS["tab_crno"])
    page.wait_for_timeout(1200)
    dismiss(page)
    if not wait_for(page, lambda: has_id(page, IDS["input_crno"])):
        raise RuntimeError("등록번호 입력창을 찾지 못함")


def set_crno(page, num):
    got = page.evaluate(SET_JS, {"id": IDS["input_crno"], "num": num})
    if re.sub(r"\D", "", got or "") != num:      # 값이 안 들어갔으면 실제 키 입력으로 재시도
        page.focus(f'[id="{IDS["input_crno"]}"]')
        page.keyboard.type(num, delay=40)
        got = page.input_value(f'[id="{IDS["input_crno"]}"]')
    return re.sub(r"\D", "", got or "") == num


def choose_row(rows, pick_first=False):
    """결과 행 목록 -> (인덱스 or None, 상태, 메모). 살아있는 등기 > 폐쇄여부 N > 본점 > 폐쇄 아님 순으로 좁힌다."""
    if not rows:
        return None, "NONE", ""
    if len(rows) == 1:
        return 0, "OK", ""
    idxs = list(range(len(rows)))

    def narrow(pred):
        sub = [i for i in idxs if pred(rows[i])]
        return sub or idxs

    idxs = narrow(lambda r: "살아있는 등기" in r["text"])
    idxs = narrow(lambda r: len(r["cells"]) >= 2 and r["cells"][-2] == "N")
    idxs = narrow(lambda r: "본점" in r["text"])
    idxs = narrow(lambda r: "폐쇄" not in r["text"])
    if len(idxs) == 1 or pick_first:
        return idxs[0], "OK", ""
    return None, "MULTI", " | ".join(rows[i]["text"][:60] for i in idxs[:5])


def read_cart_count(page):
    return page.evaluate(r"() => { const m = document.body.innerText.match(/전체\s*(\d+)\s*건/); return m ? parseInt(m[1]) : -1; }")


def apply_one(page, company, crno, opt, prev_count, rep=""):
    tag = re.sub(r"\D", "", crno)
    go_to_search(page)
    before = page.evaluate(ROWS_JS)
    if not set_crno(page, tag):
        raise RuntimeError("등록번호 입력값 확인 실패")
    dismiss(page)
    LAST_ALERTS.clear()
    if not js_click(page, IDS["btn_search"]):
        raise RuntimeError("검색 버튼을 찾지 못함")

    def probe():
        if any(re.search(r"없습니다|존재하지|올바르지", a) for a in LAST_ALERTS):
            return "NONE"
        rows_ = page.evaluate(ROWS_JS)
        return rows_ if rows_ and rows_ != before else None   # 이전 검색의 잔여 결과는 무시

    rows = wait_for(page, probe, timeout=10000)
    if rows is None or rows == "NONE":
        dump(page, f"{tag}_00_noresult")
        return ST_NOROW, (LAST_ALERTS[-1][:80] if LAST_ALERTS else "검색결과 없음(또는 응답 지연)"), "", prev_count
    page.wait_for_timeout(500)
    dump(page, f"{tag}_01_results")

    idx, st, note = choose_row(rows, opt["pick_first"])
    if st == "MULTI":
        return ST_MULTI, note, "", prev_count
    picked = rows[idx]["text"]
    soft = "" if norm_company(company) in norm_company(picked) else "(행에서 상호 미확인) "
    if opt["dry_run"]:
        return ST_DRY, soft + f"행 {len(rows)}개 중 선택예정", picked, prev_count
    if not page.evaluate(PICK_JS, idx):
        raise RuntimeError("결과 행 선택 실패")
    page.wait_for_timeout(300)

    # 화면 상태를 보고 다음 동작을 정하는 루프 (고정 순서에 의존하지 않음)
    items = opt["items"]
    state_arg = {"scope": SCOPE_TEXT, "kind": IDS["kind_malso"], "pay": IDS["btn_pay"], "add": IDS["btn_add"],
                 "item0": IDS["item_chk"].format(i=items[0]), "next": IDS["btn_next"]}
    scope_done = labels_printed = False
    rep_note = ""
    noselect_seen = NOSELECT[0]
    stt = {}
    for step in range(16):
        dismiss(page)
        if NOSELECT[0] > noselect_seen:             # '선택한 사항이 없습니다' -> 선택이 풀렸다는 뜻
            noselect_seen = NOSELECT[0]
            st_now = page.evaluate(STATE_JS, state_arg)
            if NOSELECT[0] > 6 or not (st_now["hasItems"] or st_now["hasSelAll"]):
                dump(page, f"{tag}_NOSELECT")
                raise RuntimeError("'선택한 사항이 없습니다' 안내창이 항목 화면이 아닌 곳에서 발생(전체선택 등 다른 선택이 필요한 화면일 수 있음)")
        page.wait_for_timeout(300)
        stt = page.evaluate(STATE_JS, state_arg)

        if stt["hasPay"]:                           # 완료 판정은 btn_pay 가 보일 때만 (btn_new_add 는 중간 화면에도 있을 수 있음)
            count = read_cart_count(page)
            dump(page, f"{tag}_04_pay_list")
            warn = ""
            if prev_count >= 0 and count >= 0 and count != prev_count + 1:
                warn = f"(주의: 건수 {prev_count}->{count}, 증가 미확인) "
            return ST_DONE, (soft + rep_note + warn + f"장바구니 {count}건").strip(), picked, count

        if not scope_done:
            if opt["scope"] == "partial" and stt["hasScope"]:
                page.evaluate(CLICK_TEXT_JS, SCOPE_TEXT)
                scope_done = True
                dump(page, f"{tag}_02_scope")
                page.wait_for_timeout(400)
            elif opt["scope"] == "malso" and stt["hasKind"]:
                page.evaluate("id => { const l = document.querySelector('label[for=\"' + id + '\"]'); if (l) l.click(); }", IDS["kind_malso"])
                scope_done = True
                page.wait_for_timeout(400)

        if stt["hasItems"]:
            page.wait_for_timeout(600)              # 항목 표가 다시 그려지며 체크가 풀리는 경우 대비
            if opt["list_items"]:
                rows_txt = page.evaluate(ITEM_ROWS_JS, IDS["item_grid"])
                print("    ---- 항목 목록 (행 번호: 내용) ----")
                for r_, t_ in rows_txt:
                    print(f"    행 {r_}: {t_[:90]}")
                dump(page, f"{tag}_03_items")
                return "항목목록확인", "항목 목록을 터미널에 출력함(신청 안 함)", picked, prev_count
            res = page.evaluate(CHECK_ITEMS_JS, {"rows": items, "tpl": IDS["item_chk"]})
            for i, ok in res:                       # 체크가 안 먹으면 라벨 클릭 -> 실제 마우스 클릭 순으로 재시도
                if ok is not True and ok != "missing":
                    page.evaluate(LABEL_CLICK_JS, {"grid": IDS["item_grid"], "i": i})
            res = page.evaluate(CHECK_ITEMS_JS, {"rows": items, "tpl": IDS["item_chk"]})
            for i, ok in res:
                if ok is not True:
                    try:
                        page.locator(f'[id="{IDS["item_grid"]}_cell_{i}_0"] .w2checkbox_label').click(timeout=3000, force=True)
                    except Exception:
                        pass
            res = page.evaluate(CHECK_ITEMS_JS, {"rows": items, "tpl": IDS["item_chk"]})
            bad = [i for i, ok in res if ok is not True]
            if bad:
                raise RuntimeError(f"항목 행 체크 실패: {bad}")
            if not labels_printed:
                labels_printed = True
                for i in items:
                    lab = page.evaluate(ITEM_ROWS_JS, IDS["item_grid"])
                    txt = next((t for r_, t in lab if r_ == i), "?")
                    print(f"    체크한 항목: 행 {i} = {txt[:80]}")
                dump(page, f"{tag}_03_items")
            page.wait_for_timeout(300)

        if stt["hasSelAll"] and not stt["hasItems"]:        # 임원 이름 목록 화면: 헤더의 '전체선택' 체크
            page.wait_for_timeout(500)
            how = select_all_officers(page)
            if not how:
                dump(page, f"{tag}_SELALL_FAIL")
                raise RuntimeError("임원 '전체선택' 체크 실패(직책 왼쪽 헤더 체크박스를 찾지 못함)")
            if how != "이미 체크됨":
                print(f"    임원 전체선택 체크 [{how}]")
            if not rep_note:
                body = page.locator("body").inner_text()
                if str(rep or "").strip():
                    rep_note = "(임원목록에서 대표자 확인) " if name_in_text(rep, body) else "(주의: 임원목록에서 대표자명 미발견) "
                dump(page, f"{tag}_03b_officers")
            page.wait_for_timeout(300)

        if stt["hasNext"]:
            js_click(page, IDS["btn_next"])
            page.wait_for_timeout(1800)
            dismiss(page)
            page.wait_for_timeout(500)
            dismiss(page)
        else:
            page.wait_for_timeout(1000)
    dump(page, f"{tag}_NOT_FINISHED")
    raise RuntimeError(f"신청 단계가 끝나지 않음(btn_pay 미도달; btn_new_add 있음={stt.get('hasAdd')})")


CONFIRM_PAY_JS = r"""() => {
  const vis = e => { const r = e.getBoundingClientRect(); const s = getComputedStyle(e); return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none'; };
  const label = b => (b.innerText || b.value || '').trim();
  const btns = [...document.querySelectorAll('a[id*="message_popup"], button[id*="message_popup"], input[id*="message_popup"]')].filter(vis);
  const textOf = b => { let e = b; for (let i = 0; i < 9; i++) { e = e.parentElement; if (!e) break; const t = (e.innerText || '').replace(/\s+/g, ' ').trim(); if (t.length >= 12) return t; } return ''; };
  for (const b of btns) { if (/btn_confirm/.test(b.id) && label(b) === '확인' && /결제할/.test(textOf(b))) { b.click(); return true; } }
  return false; }"""


def open_pay_list(page):
    """결제대상 목록으로 이동. 안내창('결제를 원하는 경우 [확인]')의 [확인]은 결제 화면으로 이동할 뿐 결제는 하지 않는다.
    반환: 건수(-1이면 읽기 실패) / 이동 실패 시 None"""
    try:
        dismiss(page)
        js_click(page, IDS["menu"])
        page.wait_for_timeout(2000)
        moved = page.evaluate(CONFIRM_PAY_JS)          # [확인] = 결제 화면(목록)으로 이동
        if not moved:
            js_click(page, IDS["btn_pay_list"])        # 안내창이 없으면 목록 버튼으로 시도
        page.wait_for_timeout(3000)
        dump(page, "PAYLIST")
        return read_cart_count(page)
    except Exception:
        return None


# ============================ 메인 ============================
def main():
    global DEBUG
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", required=True)
    ap.add_argument("--output")
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--dry-run", action="store_true", help="행 선택까지만 확인하고 신청은 하지 않음")
    ap.add_argument("--resume", action="store_true")
    ap.add_argument("--debug", action="store_true", help="단계별 스크린샷/텍스트 저장")
    ap.add_argument("--scope", choices=["partial", "malso"], default="partial",
                    help="partial: '일부' 선택(녹화 기준, 기본) / malso: '말소사항 포함' 선택")
    ap.add_argument("--items", default="", help="체크할 항목 행 번호(쉼표). 기본: partial=12, malso=14,15")
    ap.add_argument("--list-items", action="store_true", help="항목 목록(행 번호: 내용)만 출력하고 신청은 하지 않음(첫 1건)")
    ap.add_argument("--pick-first", action="store_true", help="결과 행이 여럿이면 자동으로 첫 후보 선택(기본은 수동확인으로 표시)")
    ap.add_argument("--batch", type=int, default=0, help="N건 담을 때마다 멈추고 결제 안내(0=멈추지 않음)")
    ap.add_argument("--include-unconfirmed", action="store_true", help="판정상태가 '확정'이 아닌 행도 등록번호가 있으면 처리")
    ap.add_argument("--url", default="https://www.iros.go.kr/index.jsp", help=argparse.SUPPRESS)
    ap.add_argument("--headless", action="store_true", help=argparse.SUPPRESS)
    ap.add_argument("--no-login-wait", action="store_true", help=argparse.SUPPRESS)
    args = ap.parse_args()
    DEBUG = args.debug
    items = [int(x) for x in re.split(r"[,\s]+", args.items) if x.strip().isdigit()] or SCOPE_DEFAULT_ITEMS[args.scope]
    opt = {"dry_run": args.dry_run, "pick_first": args.pick_first, "scope": args.scope, "items": items, "list_items": args.list_items}

    print(f"iros_apply {VERSION}")
    print(f"입력 파일: {args.input}")
    df = pd.read_excel(args.input, dtype=str).fillna("")
    if "법인등록번호" not in df.columns:
        sys.exit("입력 엑셀에 '법인등록번호' 열이 없습니다. crno_finder.py 결과 파일(…_등록번호.xlsx)을 사용하세요.")
    c_name = pick_company_column(df)
    c_rep = pick_optional_column(df, REP_ALIASES)
    print(f"기업명 열: '{c_name}' / 대표자 열: '{c_rep or '(없음)'}' / 범위: {args.scope} / 체크 항목 행: {items}")

    out_path = Path(args.output) if args.output else Path(args.input).with_name(Path(args.input).stem + "_신청결과.xlsx")
    if args.resume and out_path.exists():
        prev = pd.read_excel(out_path, dtype=str).fillna("")
        if len(prev) == len(df):
            df = prev
    for col in ("신청상태", "신청비고", "선택된행", "장바구니건수"):
        if col not in df.columns:
            df[col] = ""

    def eligible(i):
        if not re.sub(r"\D", "", df.at[i, "법인등록번호"]):
            return False
        if "판정상태" in df.columns and not args.include_unconfirmed and df.at[i, "판정상태"] != "확정":
            return False
        return df.at[i, "신청상태"] != ST_DONE

    targets = [i for i in df.index if eligible(i)]
    if args.list_items:
        targets = targets[:1]
    if args.limit:
        targets = targets[: args.limit]
    print(f"처리 대상 {len(targets)}건 / 전체 {len(df)}행")
    if not targets:
        n_crno = sum(bool(re.sub(r"\D", "", v)) for v in df["법인등록번호"])
        n_ok = int((df["판정상태"] == "확정").sum()) if "판정상태" in df.columns else -1
        print("\n[종료] 처리할 행이 없어 브라우저를 열지 않고 끝냅니다.")
        print(f"  - 법인등록번호가 채워진 행: {n_crno}개 / 판정상태 '확정': {n_ok if n_ok >= 0 else '(열 없음)'}개 / 이미 신청 완료: {int((df['신청상태'] == ST_DONE).sum())}개")
        print("  -> 입력 파일이 crno_finder.py 결과(…_등록번호.xlsx)인지 확인하세요.")
        return

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=args.headless, slow_mo=50, args=["--window-size=1400,900"])
        ctx = browser.new_context(viewport={"width": 1400, "height": 900}, locale="ko-KR")
        page = ctx.new_page()
        page.set_default_timeout(15000)

        def on_dialog(d):
            print(f"    (알림창/{d.type}) {d.message[:80]}")
            if d.type == "confirm" and "결제" in d.message:
                d.dismiss()      # 결제 진행 확인창은 '취소'
            else:
                d.accept()
        page.on("dialog", on_dialog)
        page.goto(args.url, wait_until="domcontentloaded", timeout=30000)
        page.wait_for_timeout(2500)
        close_extra_pages(ctx, page)      # 공지 팝업(별도 창) 닫기

        if not args.no_login_wait:
            print("\n[대기] 브라우저에서 로그인(아이디/자동입력방지문자)을 직접 완료하세요.")
            input("      로그인이 끝나면 Enter ▶ ")
            close_extra_pages(ctx, page)

        seen, last_count, added = set(), -1, 0
        for n, i in enumerate(targets, 1):
            company, crno = df.at[i, c_name].strip(), re.sub(r"\D", "", df.at[i, "법인등록번호"])
            if crno in seen:
                df.at[i, "신청상태"], df.at[i, "신청비고"] = "중복(건너뜀)", "같은 등록번호가 앞 행에 있음"
                continue
            seen.add(crno)
            print(f"[{n}/{len(targets)}] {company} ({crno})")
            st = note = picked = ""
            for attempt in (1, 2):            # 일시 오류는 1회 재시도
                try:
                    close_extra_pages(ctx, page)
                    st, note, picked, cnt = apply_one(page, company, crno, opt, last_count,
                                                      df.at[i, c_rep].strip() if c_rep else "")
                    if st == ST_DONE and cnt >= 0:
                        last_count = cnt
                        df.at[i, "장바구니건수"] = str(cnt)
                    break
                except Exception as e:
                    st, note, picked = ST_ERR, str(e)[:200], ""
                    dismiss(page)
                    if attempt == 1:
                        print(f"    (오류, 재시도) {note[:80]}")
                        page.wait_for_timeout(1000)
                    else:
                        traceback.print_exc()
                        try:
                            DEBUG_DIR.mkdir(parents=True, exist_ok=True)
                            page.screenshot(path=str(DEBUG_DIR / f"{crno}_ERROR.png"), full_page=True)
                        except Exception:
                            pass
            df.at[i, "신청상태"], df.at[i, "신청비고"], df.at[i, "선택된행"] = st, note, picked
            print(f"    -> {st} {note}")
            out_path.parent.mkdir(parents=True, exist_ok=True)
            df.to_excel(out_path, index=False)

            if st == ST_DONE:
                added += 1
                if args.batch and added % args.batch == 0 and not args.no_login_wait:
                    cnt = open_pay_list(page)
                    print(f"\n[일시정지] 이번 실행에서 {added}건 담음 (결제대상 목록: {cnt if cnt is not None else '확인불가'}건)")
                    input("      결제대상 목록에서 직접 결제한 뒤 Enter ▶ ")
                    last_count = -1          # 결제 후 건수가 바뀌므로 증가 검증 초기화

        print(f"\n결과 파일: {out_path}")
        if not args.dry_run and not args.list_items and not args.no_login_wait:
            cnt = open_pay_list(page)
            if cnt is None or cnt <= 0:
                print("결제대상 목록 건수를 읽지 못했습니다. 브라우저 화면에서 직접 확인하세요. (output\\debug\\PAYLIST.* 참고: --debug)")
            else:
                print(f"결제대상 목록: {cnt}건")
            print("※ 법인 결제는 페이지당 최대 10건씩 가능합니다. 결제는 직접 진행하세요.")
            input("      결제를 마친 뒤 Enter ▶ (브라우저가 닫힙니다) ")
        browser.close()


if __name__ == "__main__":
    main()
