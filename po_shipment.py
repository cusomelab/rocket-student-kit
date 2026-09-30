# -*- coding: utf-8 -*-
"""공급사 허브: 미확정 발주 조회 → 발주확정 → 쉽먼트 엑셀 업로드.

같은 Chrome(포트 9334, chrome_profile)을 쓴다.
로그인 버튼은 누르지 않는다. Access Denied면 자동 새로고침도 하지 않는다.

사용:
  python po_shipment.py list
  python po_shipment.py confirm
  python po_shipment.py upload
  python po_shipment.py pipeline
"""
from __future__ import annotations

import datetime as dt
import json
import re
import shutil
import subprocess
import sys
import time
from pathlib import Path

from playwright.sync_api import sync_playwright

sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, str(Path(__file__).parent))
import app_settings  # noqa: E402
import 쿠팡수집 as hub  # noqa: E402

BASE = Path(__file__).parent
PO_URL = "https://supplier.coupang.com/po-web/purchase/order/list"
ASN_URL = "https://supplier.coupang.com/ibs/asn/active"
ASN_BULK_DL = "https://supplier.coupang.com/ibs/shipment/parcel/bulk-creation/download"
HOME_URL = "https://supplier.coupang.com/"
PENDING_PATH = BASE / "po_pending.json"
API_DUMP = BASE / "po_apis.json"
SNAP = BASE / "po_debug.png"
CONFIRM_DL = BASE / "_po_confirm_dl"

CONFIRM_BTN = ("발주확정", "발주 확정", "공급사 확인", "공급사확인", "업체 확인", "업체확인")
PENDING_WORDS = ("미확정", "미확인", "발주등록", "발주 대기", "대기", "신규")
SKIP_CLICK = ("로그인", "가입", "아이디 찾기", "비밀번호 찾기")


def log(*a):
    print(*a, flush=True)


def snap(page, label=""):
    try:
        page.screenshot(path=str(SNAP))
        if label:
            log(f"스크린샷 ({label}): {SNAP}")
    except Exception:
        pass


def write_json(path: Path, obj) -> None:
    path.write_text(json.dumps(obj, ensure_ascii=False, indent=2), encoding="utf-8")


def wait_login(page, ctx):
    """로그인될 때까지 기다린다. 버튼은 누르지 않는다."""
    warned_manual = False
    warned_denied = False
    for i in range(720):
        if page.is_closed():
            page = ctx.new_page()
            try:
                page.goto(PO_URL, wait_until="domcontentloaded")
            except Exception:
                pass
            time.sleep(2)
        url = ""
        try:
            url = page.url or ""
        except Exception:
            url = ""
        if hub.is_access_denied(page):
            if not warned_denied:
                warned_denied = True
                log(">>> Access Denied. 자동 클릭/새로고침은 하지 않습니다.")
                log("    Chrome 주소창을 누르고 F5 한 뒤, 직접 로그인하세요.")
            time.sleep(5)
            continue
        if "xauth" in url or "login" in url.lower():
            if not warned_manual:
                warned_manual = True
                log(">>> 쿠팡 로그인 페이지입니다. Chrome에서 직접 로그인해 주세요.")
                log("    (자동으로 버튼을 누르지 않습니다.)")
            time.sleep(5)
            continue
        if "supplier.coupang.com" in url:
            time.sleep(1.5)
            return page
        if i == 0:
            log(">>> 로그인 세션을 확인합니다.")
        if i % 12 == 0:
            try:
                page.goto(PO_URL, wait_until="domcontentloaded")
            except Exception:
                pass
        time.sleep(5)
    log("RESULT: 로그인을 확인하지 못했습니다.")
    sys.exit(2)


def click_first(page, texts, timeout=2500) -> str | None:
    for t in texts:
        if t in SKIP_CLICK:
            continue
        try:
            loc = page.get_by_text(t, exact=True)
            n = loc.count()
            for i in range(n):
                el = loc.nth(i)
                if el.is_visible():
                    el.click(timeout=timeout)
                    log(f"클릭: {t}")
                    return t
        except Exception:
            pass
        try:
            loc = page.get_by_role("button", name=t)
            if loc.count() and loc.first.is_visible():
                loc.first.click(timeout=timeout)
                log(f"클릭(버튼): {t}")
                return t
        except Exception:
            pass
    return None


def require_edd() -> str:
    day = app_settings.inbound_date()
    if not day:
        log("RESULT: 대시보드에서 입고예정일을 먼저 선택하세요.")
        sys.exit(3)
    try:
        dt.datetime.strptime(day, "%Y-%m-%d")
    except ValueError:
        log("RESULT: 입고예정일 형식이 아닙니다:", day)
        sys.exit(3)
    return day


def open_hub(page, ctx=None):
    """발주 목록 페이지로 반드시 이동한다."""
    try:
        page.goto(PO_URL, wait_until="domcontentloaded", timeout=45000)
    except Exception as e:
        log(f"발주 목록 이동 실패: {e}")
        if ctx is not None:
            page = ctx.new_page()
            page.goto(PO_URL, wait_until="domcontentloaded", timeout=45000)
    time.sleep(2.5)
    log("현재 주소:", (page.url or "")[:160])
    loc = page.locator('input[placeholder="Start date"]')
    try:
        loc.wait_for(state="visible", timeout=25000)
        return page
    except Exception:
        log("날짜칸이 안 보여 새 탭으로 발주 목록을 엽니다.")
        snap(page, "no-start-date")
        if ctx is None:
            raise
        page = ctx.new_page()
        page.goto(PO_URL, wait_until="domcontentloaded", timeout=45000)
        time.sleep(2.5)
        page.locator('input[placeholder="Start date"]').wait_for(state="visible", timeout=25000)
        return page


def ensure_edd_period_type(page):
    """기간검색 종류를 입고예정일로 맞춘다."""
    try:
        txt = page.evaluate(
            """() => {
              const row = [...document.querySelectorAll('div')].find(e => {
                const t = (e.innerText || '').trim();
                return t.startsWith('기간검색') && t.length < 90;
              });
              return row ? row.innerText : '';
            }"""
        ) or ""
        if "입고예정일" in txt:
            log("기간검색 종류: 입고예정일")
            return
    except Exception:
        pass
    click_first(page, ("입고예정일",))


def _type_date(page, placeholder: str, date_str: str) -> str:
    loc = page.locator(f'input[placeholder="{placeholder}"]')
    loc.wait_for(state="visible", timeout=12000)
    loc.click()
    loc.press("Control+A")
    loc.type(date_str, delay=25)
    loc.press("Tab")
    time.sleep(0.15)
    return (loc.input_value() or "").strip()


def fill_edd_range(page, date_str: str) -> bool:
    page.locator('input[placeholder="Start date"]').wait_for(state="visible", timeout=15000)
    start = _type_date(page, "Start date", date_str)
    end = _type_date(page, "End date", date_str)
    log(f"날짜 입력: {start} ~ {end}")
    if start == date_str and end == date_str:
        return True
    snap(page, "date-mismatch")
    log("날짜가 설정값과 다릅니다. 화면을 확인하세요.")
    return False


def apply_edd_search(page, date_str: str, ctx=None):
    log(f"발주 목록에서 입고예정일 {date_str} ~ {date_str} 검색")
    page = open_hub(page, ctx or page.context)
    ensure_edd_period_type(page)
    time.sleep(0.3)
    if not fill_edd_range(page, date_str):
        log("날짜 입력에 실패했습니다.")
    time.sleep(0.2)
    btn = page.get_by_role("button", name="검색")
    if btn.count():
        btn.first.click()
        log("검색 클릭")
    else:
        clicked = click_first(page, ("검색", "조회", "Search"))
        if not clicked:
            page.keyboard.press("Enter")
            log("검색 버튼을 못 찾아 Enter로 검색합니다.")
    time.sleep(2.2)
    check = page.evaluate(
        """() => [...document.querySelectorAll('input[placeholder$="date"]')]
             .map(e => e.placeholder + ':' + e.value)"""
    )
    log("검색 후 날짜:", check)
    log("검색 클릭 후 목록을 읽습니다.")
    return page


def collect_xhr(page):
    bag = []

    def on_resp(resp):
        try:
            ct = (resp.headers or {}).get("content-type", "")
            if "json" not in ct:
                return
            if "coupang.com" not in resp.url:
                return
            data = resp.json()
            bag.append({"url": resp.url[:300], "method": resp.request.method, "data": data})
        except Exception:
            pass

    page.on("response", on_resp)
    return bag


def extract_rows_from_json(obj, acc=None, depth=0):
    if acc is None:
        acc = []
    if depth > 8:
        return acc
    if isinstance(obj, list) and obj and isinstance(obj[0], dict):
        keys = {str(k).lower() for k in obj[0].keys()}
        if keys & {"skuid", "sku_id", "sku", "orderid", "ordernumber", "purchaseordernumber", "po"}:
            for it in obj:
                if isinstance(it, dict):
                    acc.append(normalize_item(it))
            return acc
    if isinstance(obj, dict):
        for v in obj.values():
            extract_rows_from_json(v, acc, depth + 1)
    elif isinstance(obj, list):
        for v in obj[:50]:
            extract_rows_from_json(v, acc, depth + 1)
    return acc


def pick(d: dict, *names):
    lower = {str(k).lower(): v for k, v in d.items()}
    for n in names:
        if n.lower() in lower and lower[n.lower()] not in (None, ""):
            return str(lower[n.lower()]).strip()
    for k, v in d.items():
        lk = str(k).lower()
        for n in names:
            if n.lower() in lk and v not in (None, ""):
                return str(v).strip()
    return ""


def normalize_item(it: dict) -> dict:
    return {
        "orderId": pick(it, "orderId", "orderNumber", "purchaseOrderNumber", "poNumber", "po"),
        "sku": pick(it, "skuId", "sku", "skuID"),
        "name": pick(it, "skuName", "productName", "name", "itemName"),
        "center": pick(it, "center", "fc", "warehouse", "fulfillmentCenter", "centerName"),
        "qty": pick(it, "qty", "quantity", "orderQty", "confirmQty"),
        "edd": pick(it, "edd", "inboundDate", "expectedDate", "deliveryDate"),
        "status": pick(it, "status", "orderStatus", "confirmStatus", "state"),
        "raw": {k: it[k] for k in list(it)[:12]},
    }


SCRAPE_JS = """() => {
  const out = { headers: [], rows: [], buttons: [] };
  const btns = [...document.querySelectorAll('button, a, [role=button]')]
    .map(e => (e.innerText||'').trim().replace(/\\s+/g,' '))
    .filter(t => t && t.length < 30);
  out.buttons = [...new Set(btns)].slice(0, 60);
  const table = document.querySelector('table');
  if (!table) return out;
  out.headers = [...table.querySelectorAll('thead th, tr:first-child th, tr:first-child td')]
    .map(e => (e.innerText||'').trim().replace(/\\s+/g,' '));
  const bodyRows = table.querySelectorAll('tbody tr');
  const rows = bodyRows.length ? bodyRows : [...table.querySelectorAll('tr')].slice(1);
  rows.forEach(tr => {
    const cells = [...tr.querySelectorAll('td')].map(td => (td.innerText||'').trim().replace(/\\s+/g,' '));
    if (cells.filter(Boolean).length >= 2) out.rows.push(cells);
  });
  return out;
}"""


def map_scraped(headers, cells) -> dict:
    h = [x.replace(" ", "") for x in headers]
    def col(*cands):
        for c in cands:
            for i, name in enumerate(h):
                if c in name and i < len(cells):
                    return cells[i]
        return ""
    return {
        "orderId": col("발주번호", "발주서", "PO", "주문번호"),
        "sku": col("SKUID", "SKU", "Sku"),
        "name": col("상품명", "SKU명", "상품"),
        "center": col("물류센터", "센터", "납품센터", "FC"),
        "qty": col("수량", "납품수량", "발주수량"),
        "edd": col("입고예정", "예정일", "납기"),
        "status": col("상태", "진행", "확정"),
        "raw": {"cells": cells[:12]},
    }


def is_pending(row: dict) -> bool:
    st = (row.get("status") or "")
    if not st:
        return True
    if any(w in st for w in ("완료", "입고", "확정", "확인완료", "발송")):
        return False
    return any(w in st for w in PENDING_WORDS) or "확인" not in st


def scrape_orders(page, bag=None) -> list[dict]:
    rows = []
    for item in bag or []:
        rows.extend(extract_rows_from_json(item.get("data")))
    if bag:
        write_json(API_DUMP, [{"url": x["url"], "method": x["method"]} for x in bag[-40:]])
        log(f"XHR {len(bag)}건 캡처, JSON에서 주문 {len(rows)}건")

    scraped = page.evaluate(SCRAPE_JS) or {}
    log("화면 버튼:", ", ".join((scraped.get("buttons") or [])[:20]) or "(없음)")
    headers = scraped.get("headers") or []
    if not rows:
        for cells in scraped.get("rows") or []:
            rows.append(map_scraped(headers, cells))
        log(f"테이블에서 {len(rows)}건")

    seen = set()
    uniq = []
    for r in rows:
        key = (r.get("orderId"), r.get("sku"), tuple((r.get("raw") or {}).get("cells") or [])[:6])
        if key in seen:
            continue
        seen.add(key)
        uniq.append(r)
    return uniq


def list_orders(page) -> list[dict]:
    day = require_edd()
    bag = collect_xhr(page)
    page = apply_edd_search(page, day)
    return scrape_orders(page, bag)


def _accept_dialogs(page):
    def on_dialog(d):
        msg = d.message or ""
        log("알림:", msg[:300].replace("\n", " "))
        try:
            d.accept()
            log("알림 확인 클릭")
        except Exception:
            pass

    try:
        page.on("dialog", on_dialog)
    except Exception:
        pass


def select_all_rows(page) -> int:
    boxes = page.locator("input[type=checkbox]")
    if boxes.count() < 2:
        log("이 페이지에 선택할 발주가 없습니다.")
        return 0
    header = boxes.first
    header.scroll_into_view_if_needed()
    time.sleep(0.2)
    try:
        if not header.is_checked():
            header.click(timeout=3000)
            time.sleep(0.4)
    except Exception:
        header.click(timeout=3000)
        time.sleep(0.4)
    checked = page.evaluate(
        "() => [...document.querySelectorAll('input[type=checkbox]')].filter(e => e.checked).length"
    )
    log(f"전체 선택: 체크 {checked}개")
    return int(checked or 0)


def download_upload_template(page, page_no: int) -> Path | None:
    CONFIRM_DL.mkdir(exist_ok=True)
    btn = page.get_by_role("button", name="발주서 업로드 양식", exact=True)
    if not btn.count():
        log("발주서 업로드 양식 버튼을 찾지 못했습니다.")
        snap(page, "no-template-btn")
        return None
    try:
        with page.expect_download(timeout=90000) as dl:
            btn.first.click()
        item = dl.value
        name = item.suggested_filename or f"po_template_p{page_no}.xlsx"
        dest = CONFIRM_DL / f"p{page_no}_{name}"
        item.save_as(str(dest))
        log(f"양식 다운로드: {dest.name} ({dest.stat().st_size} bytes)")
        return dest
    except Exception as e:
        log(f"양식 다운로드 실패: {e}")
        snap(page, "template-dl-fail")
        return None


def attach_and_upload_confirm(popup, xlsx: Path) -> bool:
    popup.wait_for_load_state("domcontentloaded")
    time.sleep(1.5)
    btn = popup.get_by_role("button", name="발주확정 파일 업로드", exact=True)
    if btn.count():
        btn.first.click()
        log("발주확정 파일 업로드 클릭")
        time.sleep(1.2)
    else:
        log("발주확정 파일 업로드 버튼을 찾는 중...")
        click_first(popup, ("발주확정 파일 업로드",))
        time.sleep(1.2)

    agree = popup.get_by_text("위 사항에 모두 동의합니다", exact=False)
    if agree.count():
        agree.last.click()
        log("최종 동의 체크")
        time.sleep(0.3)
    else:
        boxes = popup.locator("input[type=checkbox]")
        if boxes.count():
            boxes.last.check(force=True)
            log("최종 동의 체크(체크박스)")

    attached = False
    try:
        inp = popup.locator("input[type=file]")
        if inp.count():
            inp.first.set_input_files(str(xlsx))
            attached = True
            log(f"파일 첨부: {xlsx.name}")
    except Exception as e:
        log(f"input file 실패: {e}")
    if not attached:
        try:
            with popup.expect_file_chooser(timeout=8000) as fc:
                add = popup.get_by_text("파일을 추가해주세요", exact=False)
                if add.count():
                    add.last.click()
                else:
                    popup.get_by_text("파일 첨부", exact=False).first.click()
            fc.value.set_files(str(xlsx))
            attached = True
            log(f"파일 첨부(chooser): {xlsx.name}")
        except Exception as e:
            log(f"파일 첨부 실패: {e}")
            snap(popup, "attach-fail")
            return False

    time.sleep(0.6)
    up = popup.get_by_role("button", name="업로드하기", exact=True)
    if not up.count():
        up = popup.get_by_role("button", name="업로드", exact=True)
    if up.count():
        up.last.click()
        log("업로드하기 클릭")
    else:
        click_first(popup, ("업로드하기", "업로드"))
    time.sleep(3)
    txt = ""
    try:
        txt = popup.evaluate("() => (document.body.innerText||'')") or ""
    except Exception:
        pass
    if any(k in txt for k in ("완료", "성공", "업로드되었습니다", "처리되었습니다")):
        log("업로드 완료 문구 확인")
        return True
    log("업로드 후 화면 확인 필요")
    return True


def pagination_info(page) -> dict:
    try:
        return page.evaluate(
            """() => {
              const pag = document.querySelector('.ant-pagination');
              if (!pag) return { current: 1, pages: [1], hasNext: false };
              const active = pag.querySelector('.ant-pagination-item-active');
              const current = parseInt((active && active.innerText) || '1', 10) || 1;
              const pages = [...pag.querySelectorAll('li.ant-pagination-item')]
                .map(li => parseInt((li.innerText || '').trim(), 10))
                .filter(n => n > 0);
              const next = pag.querySelector('li.ant-pagination-next');
              const hasNext = !!(next && !String(next.className).includes('disabled'));
              return { current, pages, hasNext };
            }"""
        ) or {"current": 1, "pages": [1], "hasNext": False}
    except Exception:
        return {"current": 1, "pages": [1], "hasNext": False}


def goto_page_no(page, n: int) -> bool:
    try:
        page.bring_to_front()
    except Exception:
        pass
    pag = page.locator(".ant-pagination")
    if pag.count():
        try:
            pag.first.scroll_into_view_if_needed()
        except Exception:
            pass
    item = page.locator(f"li.ant-pagination-item-{n}")
    clicked = False
    if item.count():
        item.first.click(timeout=4000)
        clicked = True
        log(f"{n}페이지 번호 클릭")
    else:
        nxt = page.locator("li.ant-pagination-next")
        cls = (nxt.first.get_attribute("class") or "") if nxt.count() else "disabled"
        if nxt.count() and "disabled" not in cls:
            try:
                nxt.locator("button").first.click(timeout=3000)
            except Exception:
                nxt.first.click(timeout=3000)
            clicked = True
            log("다음 페이지 화살표 클릭")
    if not clicked:
        return False
    try:
        page.wait_for_function(
            """(n) => {
              const a = document.querySelector('.ant-pagination-item-active');
              return a && a.innerText.trim() === String(n);
            }""",
            arg=n,
            timeout=15000,
        )
    except Exception:
        time.sleep(2)
    time.sleep(1.2)
    page.locator("input[type=checkbox]").first.wait_for(state="visible", timeout=15000)
    info = pagination_info(page)
    log(f"이동 후 현재 {info.get('current')}페이지 / 목록 {info.get('pages')}")
    return int(info.get("current") or 0) == n or True


def process_one_list_page(page, ctx, page_no: int) -> int:
    checked = select_all_rows(page)
    if checked <= 1:
        log(f"{page_no}페이지: 선택할 발주가 없습니다.")
        return 0
    xlsx = download_upload_template(page, page_no)
    if not xlsx:
        return 0
    popup = None
    try:
        with ctx.expect_page(timeout=20000) as np:
            page.get_by_role("button", name="발주서 업로드", exact=True).first.click()
        popup = np.value
        log("발주서 업로드 창:", (popup.url or "")[:120])
    except Exception as e:
        log(f"발주서 업로드 창을 못 열었습니다: {e}")
        snap(page, "no-upload-window")
        return 0
    try:
        _accept_dialogs(popup)
        ok = attach_and_upload_confirm(popup, xlsx)
        return 1 if ok else 0
    finally:
        try:
            if popup is not None and not popup.is_closed():
                time.sleep(1)
                popup.close()
        except Exception:
            pass
        try:
            page.bring_to_front()
        except Exception:
            pass


def confirm_orders(page, rows: list[dict]) -> int:
    _accept_dialogs(page)
    done = 0
    page_no = 1
    while True:
        info = pagination_info(page)
        log(
            f"=== {page_no}페이지 발주확정 "
            f"(화면 {info.get('current')} / 페이지 {info.get('pages')}) ==="
        )
        n = process_one_list_page(page, page.context, page_no)
        done += n
        log(f"{page_no}페이지 처리 끝")
        try:
            page.bring_to_front()
        except Exception:
            pass
        time.sleep(0.8)
        info = pagination_info(page)
        next_no = page_no + 1
        pages = info.get("pages") or []
        has_next = bool(info.get("hasNext")) or next_no in pages
        if not has_next:
            log("마지막 페이지입니다. 1→2→3 순서를 마쳤습니다.")
            break
        log(f"{page_no}페이지 다음 → {next_no}페이지")
        if not goto_page_no(page, next_no):
            log(f"{next_no}페이지로 이동하지 못했습니다.")
            break
        page_no = next_no
        if page_no > 40:
            log("페이지 한도를 넘었습니다.")
            break
    log(f"RESULT: 발주확정 처리 {done}페이지")
    return done


def shipment_files() -> list[Path]:
    if app_settings.is_drive():
        log(f"쉽먼트 폴더(드라이브): {app_settings.drive_url()}")
        import drive_ship
        drive_ship.list_and_download(download=True)
        files = drive_ship.local_excels()
        log(f"받은 파일 {len(files)}개")
        return files
    folder = app_settings.ship_folder()
    files = app_settings.list_ship_excels()
    log(f"쉽먼트 폴더: {folder}")
    return files


def goto_shipment_upload(page) -> bool:
    click_first(page, ("입고관리", "입고", "Inbound"))
    time.sleep(0.6)
    hit = click_first(page, (
        "쉽먼트 업로드", "쉽먼트업로드", "Shipment Upload",
        "쉽먼트 관리", "쉽먼트", "발송 정보", "발송정보입력",
    ))
    time.sleep(1.2)
    if hit:
        return True
    # URL 후보
    for url in (
        "https://supplier.coupang.com/scm/inbound/shipment",
        "https://supplier.coupang.com/scm/purchase/order/shipment",
        "https://supplier.coupang.com/scm/shipment/upload",
    ):
        try:
            page.goto(url, wait_until="domcontentloaded")
            time.sleep(2)
            if "xauth" not in (page.url or "") and "404" not in (page.title() or ""):
                body = page.evaluate("() => (document.body.innerText||'').slice(0,500)") or ""
                if "없" in body and "페이지" in body:
                    continue
                log("쉽먼트 페이지:", page.url[:120])
                return True
        except Exception:
            pass
    return False


def upload_files(page, files: list[Path]) -> int:
    if not files:
        log("업로드할 엑셀이 없습니다.")
        loc = app_settings.drive_url() if app_settings.is_drive() else str(app_settings.ship_folder())
        log(f"  설정한 폴더에 Shipment_Upload_*.xlsx 를 넣으세요: {loc}")
        return 0
    log(f"업로드 대상 {len(files)}개")
    for f in files:
        log(" -", f.name)

    if not goto_shipment_upload(page):
        log("쉽먼트 업로드 메뉴를 화면에서 찾는 중입니다.")
        snap(page, "no-shipment-menu")

    click_first(page, ("업로드", "파일 업로드", "엑셀 업로드", "Upload", "불러오기"))
    time.sleep(0.8)

    ok = 0
    for path in files:
        done = False
        for attempt in range(3):
            target = None
            try:
                loc = page.get_by_text("파일 업로드")
                if loc.count():
                    target = loc.first
            except Exception:
                pass
            if target is None:
                try:
                    loc = page.get_by_role("button", name="업로드")
                    if loc.count():
                        target = loc.first
                except Exception:
                    pass
            if target is None:
                try:
                    inp = page.locator("input[type=file]")
                    if inp.count():
                        inp.first.set_input_files(str(path))
                        done = True
                        log(f"파일 지정: {path.name}")
                        break
                except Exception as e:
                    log(f"input file 실패: {e}")
            if target is not None:
                try:
                    with page.expect_file_chooser(timeout=8000) as fc:
                        target.click()
                    fc.value.set_files(str(path))
                    done = True
                    log(f"파일 지정: {path.name}")
                    break
                except Exception as e:
                    log(f"chooser 실패({attempt}): {str(e)[:80]}")
            time.sleep(1)
        if not done:
            snap(page, "upload-fail")
            log(f"RESULT: {path.name} 업로드 창을 못 열었습니다.")
            continue
        time.sleep(1)
        click_first(page, ("업로드", "등록", "확인", "저장", "제출", "Upload"))
        time.sleep(2)
        txt = ""
        try:
            txt = page.evaluate("() => (document.body.innerText||'')") or ""
        except Exception:
            pass
        if any(k in txt for k in ("완료", "성공", "업로드됨", "complete")):
            log(f"업로드 완료: {path.name}")
        else:
            log(f"업로드 신호 대기: {path.name} (화면에서 결과를 확인하세요)")
        ok += 1
        time.sleep(1.2)
    log(f"RESULT: 쉽먼트 업로드 시도 {ok}/{len(files)}개")
    return ok


def wait_after_confirm(seconds: int = 60):
    log(f"발주확정 반영을 위해 {seconds}초 기다립니다.")
    time.sleep(seconds)


def open_shipment_bulk_search(page, date_str: str):
    """쉽먼트 일괄등록 → 양식 다운로드 → 입고예정일 검색까지."""
    log("쉽먼트 일괄등록 화면으로 이동합니다.")
    try:
        page.bring_to_front()
    except Exception:
        pass
    page.goto(ASN_URL, wait_until="domcontentloaded")
    time.sleep(3)
    bulk = page.get_by_text("쉽먼트 일괄등록", exact=True)
    if not bulk.count():
        bulk = page.locator("button", has_text="쉽먼트 일괄등록")
    if bulk.count():
        bulk.first.click()
        log("쉽먼트 일괄등록 클릭")
        time.sleep(1.2)
    else:
        log("쉽먼트 일괄등록 버튼을 못 찾아 양식 페이지로 바로 갑니다.")
        page.goto(ASN_BULK_DL, wait_until="domcontentloaded")
        time.sleep(2)

    dl = page.get_by_text("일괄등록 양식 다운로드", exact=True)
    if dl.count():
        dl.first.click()
        log("일괄등록 양식 다운로드 클릭")
        time.sleep(2.5)
    if "bulk-creation/download" not in (page.url or ""):
        page.goto(ASN_BULK_DL, wait_until="domcontentloaded")
        time.sleep(2.5)
        log("양식 다운로드 페이지로 이동:", (page.url or "")[:120])

    edd = page.locator('input[placeholder="Select EDD"]')
    edd.wait_for(state="visible", timeout=20000)
    edd.click()
    edd.press("Control+A")
    edd.type(date_str, delay=25)
    edd.press("Tab")
    time.sleep(0.3)
    got = (edd.input_value() or "").strip()
    log(f"입고예정일(EDD) 입력: {got}")
    if got != date_str:
        snap(page, "edd-mismatch")
        log("EDD가 설정값과 다릅니다.")

    search = page.locator("button.primary-button", has_text="검색")
    if not search.count():
        search = page.get_by_role("button", name="검색", exact=True)
    search.first.click()
    log("검색 클릭")
    time.sleep(2.5)
    log("RESULT: 쉽먼트 일괄등록 양식 — 입고예정일 검색까지 완료")
    download_bulk_templates_by_page(page)
    run_sheet_shipment_plan()
    return True


def bulk_page_info(page) -> dict:
    try:
        return page.evaluate(
            """() => {
              const pag = document.querySelector('ul.bootpag, ul.pagination');
              if (!pag) return { current: 1, pages: [1], hasNext: false };
              let current = 1;
              const pages = [];
              let hasNext = false;
              [...pag.querySelectorAll('li')].forEach(li => {
                const t = (li.innerText || '').trim();
                const cls = li.className || '';
                if (t === '다음') hasNext = !cls.includes('disabled');
                if (/^\\d+$/.test(t)) {
                  const n = parseInt(t, 10);
                  pages.push(n);
                  if (cls.includes('active')) current = n;
                }
              });
              return { current, pages, hasNext };
            }"""
        ) or {"current": 1, "pages": [1], "hasNext": False}
    except Exception:
        return {"current": 1, "pages": [1], "hasNext": False}


def set_po_check_all(page, checked: bool) -> int:
    box = page.locator("#check-all")
    box.wait_for(state="visible", timeout=15000)
    box.scroll_into_view_if_needed()
    time.sleep(0.2)
    for _ in range(3):
        on = False
        try:
            on = bool(box.is_checked())
        except Exception:
            on = bool(page.evaluate("() => document.querySelector('#check-all')?.checked"))
        if on == checked:
            break
        box.click()
        time.sleep(0.45)
    n = page.evaluate(
        "() => [...document.querySelectorAll('input[name=poSelected]')].filter(e => e.checked).length"
    )
    log(f"발주 선택 {'체크' if checked else '해제'}: 행 {n}개")
    return int(n or 0)


def goto_bulk_page(page, n: int) -> bool:
    pag = page.locator("ul.pagination.bootpag, ul.bootpag")
    if pag.count():
        try:
            pag.first.scroll_into_view_if_needed()
        except Exception:
            pass
    item = pag.locator("li").filter(has_text=re.compile(rf"^{n}$"))
    if item.count():
        try:
            item.locator("a").first.click(timeout=4000)
        except Exception:
            item.first.click(timeout=4000)
        log(f"{n}페이지 클릭")
    else:
        nxt = page.locator("ul.pagination.bootpag li.next:not(.disabled) a")
        if not nxt.count():
            return False
        nxt.first.click()
        log("다음 페이지 클릭")
    time.sleep(1.6)
    info = bulk_page_info(page)
    log(f"이동 후 {info.get('current')}페이지 / {info.get('pages')}")
    return True


def save_bulk_excel(src: Path) -> Path:
    dest_dir = app_settings.ship_folder()
    dest_dir.mkdir(parents=True, exist_ok=True)
    dest = dest_dir / src.name
    if src.resolve() != dest.resolve():
        shutil.copy2(src, dest)
    log(f"로컬 저장: {dest}")
    return dest


def download_bulk_templates_by_page(page) -> list[Path]:
    """한 페이지씩 발주선택 → 양식 다운로드 → 선택 해제 → 다음 페이지."""
    _accept_dialogs(page)
    CONFIRM_DL.mkdir(exist_ok=True)
    saved: list[Path] = []
    page_no = 1
    while True:
        info = bulk_page_info(page)
        log(f"=== 쉽먼트 양식 {page_no}페이지 (화면 {info.get('current')} / {info.get('pages')}) ===")
        n = set_po_check_all(page, True)
        if n <= 0:
            log(f"{page_no}페이지: 선택된 발주가 없습니다.")
            set_po_check_all(page, False)
            break
        dest = None
        try:
            with page.expect_download(timeout=120000) as dl:
                btn = page.get_by_role("button", name="양식 다운로드", exact=True)
                if not btn.count():
                    btn = page.locator("button", has_text="양식 다운로드")
                btn.first.click()
            item = dl.value
            name = item.suggested_filename or f"shipment_bulk_p{page_no}.xlsx"
            raw = CONFIRM_DL / f"p{page_no}_{name}"
            item.save_as(str(raw))
            dest = save_bulk_excel(raw)
            saved.append(dest)
            log(f"{page_no}페이지 양식 다운로드: {dest.name} ({dest.stat().st_size} bytes)")
        except Exception as e:
            log(f"{page_no}페이지 양식 다운로드 실패: {e}")
            snap(page, f"bulk-dl-p{page_no}")
        log("발주 선택을 해제한 뒤 다음 페이지로 갑니다.")
        set_po_check_all(page, False)
        time.sleep(0.4)
        info = bulk_page_info(page)
        next_no = page_no + 1
        pages = info.get("pages") or []
        has_next = bool(info.get("hasNext")) or next_no in pages
        if not has_next:
            log("마지막 페이지입니다.")
            break
        if not goto_bulk_page(page, next_no):
            log(f"{next_no}페이지로 이동하지 못했습니다.")
            break
        page_no = next_no
        if page_no > 40:
            log("페이지 한도를 넘었습니다.")
            break
    if saved and app_settings.is_drive():
        log("쉽먼트 드라이브 폴더에 올립니다.")
        try:
            import drive_ship
            drive_ship.upload_local_files(page.context, saved)
        except Exception as e:
            log(f"드라이브 업로드 오류: {e}")
    log(f"RESULT: 쉽먼트 양식 {len(saved)}개 저장")
    return saved


def run_sheet_shipment_plan() -> int:
    log("구글시트 쉽먼트 계획 ①~⑤를 순서대로 실행합니다.")
    proc = subprocess.Popen(
        [sys.executable, "-u", str(BASE / "sheet_shipment.py")],
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        encoding="utf-8",
        errors="replace",
        cwd=str(BASE),
    )
    assert proc.stdout is not None
    for line in proc.stdout:
        log(line.rstrip())
    code = proc.wait()
    if code == 0:
        log("시트 쉽먼트 계획 ①~⑤ 완료")
    else:
        log(f"시트 쉽먼트 계획 종료 코드 {code}")
    return int(code or 0)


def save_pending(rows, extra=None):
    pending = [r for r in rows if is_pending(r)]
    payload = {
        "fetched_at": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "total": len(rows),
        "pending": len(pending),
        "rows": rows[:500],
        "pending_rows": pending[:500],
        "extra": extra or {},
    }
    write_json(PENDING_PATH, payload)
    log(f"저장: {PENDING_PATH.name} (전체 {len(rows)} / 미확정 {len(pending)})")
    return payload


def existing_hub_tab(ctx):
    po = None
    any_hub = None
    for pg in ctx.pages:
        try:
            if pg.is_closed():
                continue
            url = pg.url or ""
        except Exception:
            continue
        if "xauth" in url:
            continue
        if "po-web/purchase/order" in url:
            po = pg
        elif "supplier.coupang.com" in url:
            any_hub = pg
    return po or any_hub


def run(cmd: str) -> int:
    hub.ensure_chrome()
    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp(f"http://127.0.0.1:{hub.PORT}")
        ctx = browser.contexts[0]
        page = existing_hub_tab(ctx)
        owned = page is None
        if page is None:
            page = ctx.new_page()
            try:
                page.goto(PO_URL, wait_until="domcontentloaded")
            except Exception:
                pass
            time.sleep(2)
        else:
            log("이미 열린 서플라이 허브 탭을 사용합니다.")
        try:
            page = wait_login(page, ctx)

            if cmd == "list":
                rows = list_orders(page)
                save_pending(rows)
                log("RESULT: 발주 조회 완료")
                return 0

            if cmd == "confirm":
                day = require_edd()
                bag = collect_xhr(page)
                page = apply_edd_search(page, day)
                rows = scrape_orders(page, bag)
                save_pending(rows)
                n = confirm_orders(page, rows)
                time.sleep(2)
                log("RESULT: 발주확정 완료" if n else "RESULT: 확정할 발주 없음")
                if n:
                    wait_after_confirm(60)
                    open_shipment_bulk_search(page, day)
                return 0

            if cmd == "upload":
                files = shipment_files()
                upload_files(page, files)
                return 0

            if cmd == "pipeline":
                day = require_edd()
                log("=== 1) 입고예정일 검색 ===")
                bag = collect_xhr(page)
                page = apply_edd_search(page, day)
                rows = scrape_orders(page, bag)
                save_pending(rows)
                log("=== 2) 발주확정 ===")
                n = confirm_orders(page, rows)
                time.sleep(2)
                log("=== 3) 쉽먼트 일괄등록 양식 검색 ===")
                if n:
                    wait_after_confirm(60)
                    open_shipment_bulk_search(page, day)
                log("RESULT: 발주확정 → 쉽먼트 일괄등록 검색 끝")
                return 0

            log("알 수 없는 명령:", cmd)
            return 1
        finally:
            if owned:
                try:
                    if page is not None and not page.is_closed():
                        page.close()
                except Exception:
                    pass


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else "list"
    sys.exit(run(cmd))
