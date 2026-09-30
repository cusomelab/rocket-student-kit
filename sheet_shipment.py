# -*- coding: utf-8 -*-
"""구글시트 메뉴: 🚚 쉽먼트 계획 ①~⑤를 순서대로 실행한다.

각 단계가 끝날 때까지 기다린 뒤 다음으로 넘어간다.
로그인된 CDP Chrome(포트 9334)을 사용한다.
"""
from __future__ import annotations

import re
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
DEBUG = Path(__file__).parent / "sheet_shipment_debug.png"
MENU_NAMES = ("🚚 쉽먼트 계획", "쉽먼트 계획")

STEPS = [
    {
        "label": "① 쉽먼트 리스트업",
        "items": (
            "🚀 쉽먼트 리스트업 (엑셀 가져오기)",
            "① 쉽먼트 리스트업 (엑셀 가져오기)",
            "쉽먼트 리스트업 (엑셀 가져오기)",
            "쉽먼트 리스트업",
        ),
        "ok": ("쉽먼트 데이터 수집 완료", "건 추가", "새 파일 없음"),
        "fail": ("쉽먼트 시트를 찾을 수 없습니다", "폴더 ID 입력 필요", "폴더 ID 오류"),
        "timeout": 720,
    },
    {
        "label": "② 시트 동기화",
        "items": (
            "② 시트 동기화 (입고리스트+쉽먼트)",
            "시트 동기화 (입고리스트+쉽먼트)",
        ),
        "ok": ("업데이트 완료", "건 업데이트", "동기화 완료"),
        "fail": ("시트 확인 필요", "시트 누락"),
        "timeout": 720,
    },
    {
        "label": "③ 쉽먼트 계획 생성",
        "items": (
            "③ 쉽먼트 계획 생성 (기본)",
            "쉽먼트 계획 생성 (기본)",
        ),
        "ok": ("쉽먼트 계획 생성 및 재고 배정 완료", "재고 배정 완료"),
        "fail": ("필수 시트 누락",),
        "timeout": 600,
    },
    {
        "label": "④ 쉽먼트시트 정렬",
        "items": (
            "④ 쉽먼트시트 정렬 (날짜>센터>유형>박스)",
            "쉽먼트시트 정렬 (날짜>센터>유형>박스)",
            "쉽먼트시트 정렬",
        ),
        "ok": ("정렬 완료",),
        "fail": (),
        "timeout": 300,
    },
    {
        "label": "⑤ 예상 쉽먼트 생성",
        "items": (
            "⑤ 예상 쉽먼트 생성 (I열 빈칸 전용)",
            "⑤ 예상 쉽먼트 생성 (I열 빈칸 적용)",
            "예상 쉽먼트 생성 (I열 빈칸 전용)",
            "예상 쉽먼트 생성 (I열 빈칸 적용)",
            "예상 쉽먼트 생성",
        ),
        "ok": ("이미지 노출 완료", "작업 대상 전용 이미지", "예상 쉽먼트 생성 완료"),
        "fail": ("시트 확인 불가",),
        "timeout": 720,
    },
]


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


def click_plan_item(page, names: tuple[str, ...]) -> bool:
    page.keyboard.press("Escape")
    time.sleep(0.5)
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
        log("RESULT: '쉽먼트 계획' 메뉴를 찾지 못했습니다.")
        snap(page, "no-menu")
        return False
    menu.first.click()
    time.sleep(1.0)
    for name in names:
        item = page.get_by_text(name, exact=True)
        if item.count() == 0:
            item = page.get_by_text(name)
        if item.count():
            item.first.click()
            log(f"클릭: 쉽먼트 계획 > {name}")
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


def main() -> int:
    app_settings.require_order_sheet()
    dialogs: list[str] = []
    with sync_playwright() as p:
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
        if "쉽먼트 계획" not in body:
            log("커스텀 메뉴 대기 8초...")
            time.sleep(8)

        for i, step in enumerate(STEPS, 1):
            log(f"=== {step['label']} ({i}/{len(STEPS)}) ===")
            dialogs.clear()
            if not click_plan_item(page, step["items"]):
                return 10 + i
            ok = wait_script(
                page,
                ok_needles=step["ok"],
                fail_needles=step["fail"],
                timeout_sec=step["timeout"],
                dialogs=dialogs,
            )
            fail_dlg = any(
                any(k in m for k in ("실패", "오류", "누락", "없습니다")) and "새 파일 없음" not in m
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
        log("RESULT: 쉽먼트 계획 ①~⑤ 모두 완료")
        return 0


if __name__ == "__main__":
    sys.exit(main())
