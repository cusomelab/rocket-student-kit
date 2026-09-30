# -*- coding: utf-8 -*-
"""발주서 만들기 전체 공정.

1) 서플라이 허브 SKU 목록에서 그 주 월요일~오늘(발주일) 검색
2) 상품목록 전체 다운로드
3) INBOX 드라이브 폴더에 업로드
4) 시트 로켓발주 ①~⑤ 커밋까지 순서 실행
"""
from __future__ import annotations

import datetime as dt
import subprocess
import sys
import time
from pathlib import Path

from playwright.sync_api import sync_playwright

sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, str(Path(__file__).parent))
import app_settings  # noqa: E402
import drive_ship  # noqa: E402
import 쿠팡수집 as hub  # noqa: E402

BASE = Path(__file__).parent
SKU_URL = "https://supplier.coupang.com/scm/purchase/order/sku/list"
DL_DIR = BASE / "_order_inbox_dl"
SNAP = BASE / "order_build_debug.png"


def log(*a):
    print(*a, flush=True)


def snap(page, label=""):
    try:
        page.screenshot(path=str(SNAP))
        if label:
            log(f"스크린샷 ({label}): {SNAP}")
    except Exception:
        pass


def week_range():
    today = dt.date.today()
    monday = today - dt.timedelta(days=today.weekday())
    return monday.isoformat(), today.isoformat()


def _type_named(page, name: str, value: str) -> str:
    loc = page.locator(f'input[name="{name}"]')
    loc.wait_for(state="visible", timeout=15000)
    loc.click()
    loc.press("Control+A")
    loc.type(value, delay=25)
    loc.press("Tab")
    time.sleep(0.15)
    return (loc.input_value() or "").strip()


def download_sku_list(page) -> Path | None:
    start, end = week_range()
    log(f"발주일 검색: {start} ~ {end}")
    page.goto(SKU_URL, wait_until="domcontentloaded")
    time.sleep(3)
    url = page.url or ""
    if "xauth" in url or "login" in url.lower():
        log(">>> 쿠팡 로그인 페이지입니다. Chrome에서 직접 로그인해 주세요.")
        for _ in range(120):
            time.sleep(5)
            url = page.url or ""
            if "supplier.coupang.com" in url and "xauth" not in url:
                page.goto(SKU_URL, wait_until="domcontentloaded")
                time.sleep(2)
                break
        else:
            log("RESULT: 로그인을 확인하지 못했습니다.")
            return None
    try:
        page.select_option("#searchDateType", value="PURCHASE_ORDER_DATE")
        log("기간검색: 발주일")
    except Exception as e:
        log(f"기간검색 선택 실패: {e}")
    got_s = _type_named(page, "searchStartDate", start)
    got_e = _type_named(page, "searchEndDate", end)
    log(f"날짜 입력: {got_s} ~ {got_e}")
    search = page.get_by_role("button", name="검색", exact=True)
    if not search.count():
        search = page.locator("button", has_text="검색")
    search.first.click()
    log("검색 클릭")
    time.sleep(2.5)

    def on_dialog(d):
        log("알림:", (d.message or "")[:240].replace("\n", " "))
        try:
            d.accept()
        except Exception:
            pass

    page.on("dialog", on_dialog)
    btn = page.get_by_role("button", name="상품목록 다운로드", exact=True)
    if not btn.count():
        btn = page.get_by_text("상품목록 다운로드", exact=True)
    if not btn.count():
        log("상품목록 다운로드 버튼을 찾지 못했습니다.")
        snap(page, "no-dl-btn")
        return None
    btn.first.click()
    log("상품목록 다운로드 클릭")
    try:
        page.get_by_text("선택항목 다운로드", exact=False).wait_for(state="visible", timeout=10000)
    except Exception:
        time.sleep(1.2)
    all_btn = page.get_by_text("전체 다운로드 하기", exact=True)
    if not all_btn.count():
        all_btn = page.get_by_text("전체다운로드 하기", exact=True)
    if not all_btn.count():
        all_btn = page.get_by_role("button", name="전체 다운로드 하기")
    if not all_btn.count():
        all_btn = page.locator("button", has_text="전체 다운로드")
    if not all_btn.count():
        log("전체다운로드 하기 버튼을 찾지 못했습니다.")
        snap(page, "no-all-dl")
        return None
    DL_DIR.mkdir(exist_ok=True)
    try:
        with page.expect_download(timeout=180000) as dl:
            all_btn.first.click()
            log("전체다운로드 하기 클릭")
        item = dl.value
        name = item.suggested_filename or f"sku_list_{start}_{end}.csv"
        dest = DL_DIR / name
        item.save_as(str(dest))
        log(f"다운로드: {dest.name} ({dest.stat().st_size} bytes)")
        return dest
    except Exception as e:
        log(f"다운로드 실패: {e}")
        snap(page, "dl-fail")
        return None


def run_sheet_all() -> int:
    log("시트 로켓발주 ①~커밋을 순서대로 실행합니다.")
    proc = subprocess.Popen(
        [sys.executable, "-u", str(BASE / "sheet_order.py"), "all"],
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        encoding="utf-8",
        errors="replace",
        cwd=str(BASE),
    )
    assert proc.stdout is not None
    for line in proc.stdout:
        log(line.rstrip())
    return int(proc.wait() or 0)


def main() -> int:
    app_settings.require_order_sheet()
    if not app_settings.parse_drive_id(app_settings.inbox_folder_raw()):
        log("RESULT: 설정 탭에서 'INBOX 폴더(구글 드라이브 주소)'를 먼저 저장해 주세요.")
        return 2
    start, end = week_range()
    log(f"발주서 만들기 시작 (발주일 {start} ~ {end})")
    hub.ensure_chrome()
    with sync_playwright() as p:
        b = p.chromium.connect_over_cdp(f"http://127.0.0.1:{hub.PORT}")
        ctx = b.contexts[0]
        page = ctx.new_page()
        try:
            path = download_sku_list(page)
            if not path:
                log("RESULT: 상품목록 다운로드 실패")
                return 3
            log("INBOX 드라이브에 업로드합니다.")
            inbox = app_settings.inbox_drive_url()
            log("INBOX 폴더:", inbox)
            n = drive_ship.upload_local_files(ctx, [path], folder_url=inbox)
            if n < 1:
                log("RESULT: 드라이브 업로드 실패")
                return 4
        finally:
            try:
                if not page.is_closed():
                    page.close()
            except Exception:
                pass
    time.sleep(3)
    code = run_sheet_all()
    if code == 0:
        log("RESULT: 발주서 만들기 전체 완료")
    return code


if __name__ == "__main__":
    sys.exit(main())
