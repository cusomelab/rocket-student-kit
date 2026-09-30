# -*- coding: utf-8 -*-
"""구글시트 메뉴: 🚀 로켓발주로 발주서를 만든다.

한 번에 실행: ① INBOX → ② 중복정리 → ③ 미발주집계 → ④ 재고차감 → 발주서 채우기 → ⑤ 커밋
"""
from __future__ import annotations

import sys
import time
from pathlib import Path

from playwright.sync_api import sync_playwright

sys.stdout.reconfigure(encoding="utf-8")
PORT = 9334
sys.path.insert(0, str(Path(__file__).parent))
import app_settings  # noqa: E402

SHEET = app_settings.order_sheet_url()
SHEET_ID = app_settings.order_sheet_id()
DEBUG = Path(__file__).parent / "sheet_order_debug.png"
MENU_NAMES = ("🚀 로켓발주", "로켓발주")

FILL_STEPS = [
    {
        "label": "① INBOX 가져오기",
        "items": (
            "① INBOX 가져오기 (CSV 권장)",
            "INBOX 가져오기 (CSV 권장)",
            "INBOX 가져오기",
        ),
        "ok": ("완료: 신규", "건 / 업데이트", "새 파일 없음"),
        "fail": ("INBOX_FOLDER_ID", "폴더 ID 오류"),
        "timeout": 720,
    },
    {
        "label": "② RAW 중복 정리",
        "items": (
            "② RAW 중복 정리 (최근 한 달)",
            "RAW 중복 정리 (최근 한 달)",
            "RAW 중복 정리",
        ),
        "ok": ("중복 정리 완료", "제거할 중복 데이터가 없습니다", "데이터 없음"),
        "fail": ("제목줄 오류",),
        "timeout": 600,
    },
    {
        "label": "③ 미발주 집계 갱신",
        "items": (
            "③ 미발주 집계 갱신",
            "미발주 집계 갱신",
        ),
        "ok": ("미발주 및 AC열 상품 집계 완료", "집계 완료", "미발주 피벗 갱신 완료"),
        "fail": ("필수 시트 누락",),
        "timeout": 600,
    },
    {
        "label": "④ 국내/중국 재고 차감",
        "items": (
            "④ 국내/중국 재고 우선 차감",
            "국내/중국 재고 우선 차감",
        ),
        "ok": ("재고 실시간 동기화 완료", "동기화 완료"),
        "fail": ("필수 시트 누락",),
        "timeout": 720,
    },
    {
        "label": "발주서 채우기",
        "items": (
            "④ 발주서 채우기 (세트 처리)",
            "발주서 채우기 (세트 처리)",
            "발주서 채우기",
        ),
        "ok": ("발주서 채우기 완료", "발주 대상 상품이 없습니다", "피벗 데이터 없음"),
        "fail": ("시트 확인 필요",),
        "timeout": 600,
    },
]

COMMIT_STEP = {
    "label": "⑤ 발주 커밋",
    "items": (
        "⑤ 발주 커밋(엑셀 저장/피벗정리)",
        "발주 커밋(엑셀 저장/피벗정리)",
        "발주 커밋",
    ),
    "ok": ("발주 확정 및 시트 정리 완료", "시트 정리 완료", "엑셀 발행 완료", "발주 확정 및 엑셀"),
    "fail": ("필수 시트 누락", "확정할 내역이 없습니다"),
    "timeout": 600,
}


def log(*a):
    print(*a, flush=True)


def snap(page, label=""):
    try:
        page.screenshot(path=str(DEBUG))
        if label:
            log(f"스크린샷 ({label}): {DEBUG}")
    except Exception:
        pass


def toast_text(page) -> str:
    try:
        return page.evaluate(
            """() => {
              const nodes = [...document.querySelectorAll(
                '#docs-butterbar-container, .docs-butterbar, .docs-toast, [aria-live], .script-application-sidebar'
              )];
              return nodes.map(e => (e.innerText || '').trim()).filter(Boolean).join('\\n');
            }"""
        ) or ""
    except Exception:
        return ""


def body_tail(page) -> str:
    try:
        return page.evaluate("() => (document.body.innerText || '').slice(-3000)") or ""
    except Exception:
        return ""


def dismiss_toast(page):
    try:
        page.keyboard.press("Escape")
    except Exception:
        pass
    try:
        page.locator("#docs-butterbar-container .docs-icon-close, #docs-butterbar-container [aria-label='닫기']").first.click(timeout=800)
    except Exception:
        pass
    try:
        page.evaluate(
            """() => {
              const bar = document.querySelector('#docs-butterbar-container');
              if (bar) bar.style.pointerEvents = 'none';
            }"""
        )
    except Exception:
        pass
    time.sleep(0.4)


def click_order_item(page, names: tuple[str, ...]) -> bool:
    dismiss_toast(page)
    menu = None
    for m in MENU_NAMES:
        loc = page.get_by_text(m, exact=True)
        if loc.count():
            menu = loc
            break
        loc = page.get_by_text(m)
        if loc.count():
            menu = loc
            break
    if menu is None:
        log("RESULT: '로켓발주' 메뉴를 찾지 못했습니다.")
        snap(page, "no-menu")
        return False
    try:
        menu.first.click(timeout=8000)
    except Exception:
        menu.first.click(force=True, timeout=8000)
    time.sleep(1.0)
    for name in names:
        item = page.get_by_text(name, exact=True)
        if item.count() == 0:
            item = page.get_by_text(name)
        if item.count():
            item.first.click()
            log(f"클릭: 로켓발주 > {name}")
            return True
    log(f"RESULT: 메뉴 항목을 찾지 못했습니다: {names[0]}")
    snap(page, "no-item")
    try:
        page.keyboard.press("Escape")
    except Exception:
        pass
    return False


def wait_script(page, ok_needles, fail_needles, timeout_sec, dialogs):
    t0 = time.time()
    seen_running = False
    last_toast = ""
    while time.time() - t0 < timeout_sec:
        toast = toast_text(page)
        tail = body_tail(page)
        blob = f"{toast}\n{tail}\n" + "\n".join(dialogs)
        if toast and toast != last_toast:
            log("토스트:", toast.replace("\n", " | ")[:240])
            last_toast = toast
        for n in fail_needles:
            if n and n in blob:
                log(f"RESULT: 실패 문구 감지 ({n})")
                snap(page, "fail")
                return False
        for n in ok_needles:
            if n and n in blob:
                log(f"완료 문구 감지: {n}")
                return True
        running = any(
            k in blob
            for k in (
                "스크립트 실행 중",
                "실행하고 있습니다",
                "Running script",
                "script is running",
                "스크립트를 실행",
            )
        )
        if running:
            seen_running = True
        elif seen_running:
            time.sleep(3)
            toast2 = toast_text(page) + "\n" + body_tail(page) + "\n" + "\n".join(dialogs)
            for n in ok_needles:
                if n and n in toast2:
                    log(f"완료 문구 감지: {n}")
                    return True
            log("스크립트 실행 표시가 사라졌습니다. 다음 단계로 갑니다.")
            return True
        time.sleep(1.4)
    log("RESULT: 스크립트 완료 신호를 시간 내에 못 봤습니다.")
    snap(page, "timeout")
    return False


def run_steps(page, steps, dialogs) -> int:
    for i, step in enumerate(steps, 1):
        log(f"=== {step['label']} ({i}/{len(steps)}) ===")
        dialogs.clear()
        if not click_order_item(page, step["items"]):
            return 10 + i
        ok = wait_script(
            page,
            ok_needles=step["ok"],
            fail_needles=step["fail"],
            timeout_sec=step["timeout"],
            dialogs=dialogs,
        )
        fail_dlg = any(
            any(k in m for k in ("실패", "오류", "누락")) and "없습니다" not in m
            for m in dialogs
        )
        if fail_dlg:
            log(f"RESULT: {step['label']} 알림 실패")
            return 20 + i
        if not ok:
            log(f"RESULT: {step['label']} 완료 확인 실패")
            return 20 + i
        log(f"{step['label']} 완료")
        time.sleep(2.5)
    return 0


def open_sheet(p):
    b = p.chromium.connect_over_cdp(f"http://127.0.0.1:{PORT}")
    ctx = b.contexts[0]
    page = next(
        (
            pg
            for pg in ctx.pages
            if SHEET_ID in (pg.url or "")
            and not pg.is_closed()
        ),
        None,
    )
    if page is None:
        page = ctx.new_page()
    page.goto(SHEET, wait_until="domcontentloaded")
    time.sleep(6)
    return page


def main() -> int:
    app_settings.require_order_sheet()
    cmd = sys.argv[1] if len(sys.argv) > 1 else "fill"
    dialogs: list[str] = []
    with sync_playwright() as p:
        page = open_sheet(p)
        if "accounts.google.com" in page.url or "signin" in (page.url or "").lower():
            log("RESULT: 구글 로그인이 필요합니다.")
            return 2

        def on_dialog(d):
            msg = d.message or ""
            dialogs.append(msg)
            log("시트 알림:", msg[:240])
            try:
                d.accept()
            except Exception:
                pass

        page.on("dialog", on_dialog)
        log("시트 진입:", (page.title() or "")[:60])
        body = page.evaluate("() => document.body.innerText") or ""
        if "로켓발주" not in body:
            log("커스텀 메뉴 대기 8초...")
            time.sleep(8)

        if cmd == "finish":
            code = run_steps(page, FILL_STEPS[-1:] + [COMMIT_STEP], dialogs)
            if code == 0:
                log("RESULT: 발주서 채우기·커밋 완료")
            return code

        if cmd == "commit":
            code = run_steps(page, [COMMIT_STEP], dialogs)
            if code == 0:
                log("RESULT: 발주 커밋 완료")
            return code

        if cmd == "fill":
            code = run_steps(page, FILL_STEPS, dialogs)
            if code == 0:
                log("RESULT: 발주서 채우기까지 완료. 확인 후 커밋하세요.")
            return code

        code = run_steps(page, FILL_STEPS + [COMMIT_STEP], dialogs)
        if code == 0:
            log("RESULT: 발주서 만들기 ①~커밋 모두 완료")
        return code


if __name__ == "__main__":
    sys.exit(main())
