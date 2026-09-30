# -*- coding: utf-8 -*-
"""구글 드라이브 쉽먼트 폴더에서 xlsx 목록/다운로드.

로그인된 CDP Chrome(포트 9334)을 사용한다.
"""
from __future__ import annotations

import datetime as dt
import json
import re
import sys
import time
from pathlib import Path

from playwright.sync_api import sync_playwright

sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, str(Path(__file__).parent))
import app_settings  # noqa: E402
import 쿠팡수집 as hub  # noqa: E402

BASE = Path(__file__).parent
DL_DIR = BASE / "_ship_dl"
SNAP = BASE / "drive_ship_debug.png"

LIST_JS = """() => {
  const out = [];
  const seen = new Set();
  const add = (id, name) => {
    id = (id || '').trim();
    name = (name || '').replace(/\\s+/g, ' ').trim();
    if (!id || id.length < 10 || !name || seen.has(id)) return;
    const base = name.split(',')[0].trim();
    seen.add(id);
    out.push({ id, name: base });
  };
  document.querySelectorAll('[data-id]').forEach((el) => {
    const id = el.getAttribute('data-id') || '';
    let name = el.getAttribute('aria-label') || '';
    if (!name) {
      const t = (el.innerText || '').trim().split('\\n')[0];
      name = t;
    }
    add(id, name);
  });
  document.querySelectorAll('a[href*="/file/d/"]').forEach((a) => {
    const m = (a.getAttribute('href') || '').match(/\\/file\\/d\\/([^/]+)/);
    if (m) add(m[1], a.innerText || a.getAttribute('aria-label') || '');
  });
  return out;
}"""


def log(*a):
    print(*a, flush=True)


def _save_cache(url: str, files: list[dict]) -> None:
    app_settings.DRIVE_CACHE.write_text(
        json.dumps(
            {
                "url": url,
                "fetched_at": dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "files": files,
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )


def load_cache() -> dict:
    p = app_settings.DRIVE_CACHE
    if not p.exists():
        return {}
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return {}


def _open_folder(page, url: str) -> None:
    page.goto(url, wait_until="domcontentloaded")
    time.sleep(6)


def _filter_xlsx(items: list[dict]) -> list[dict]:
    out = []
    for it in items:
        name = it.get("name") or ""
        if not name.lower().endswith(".xlsx") and "xlsx" not in name.lower():
            if not app_settings.is_ship_excel(name):
                continue
        if not app_settings.is_ship_excel(name) and "xlsx" not in name.lower():
            continue
        if name.lower() == "shipmentsupload.xlsx":
            continue
        if not app_settings.is_ship_excel(name):
            continue
        out.append(it)
    return out


def list_and_download(download: bool = False) -> list[dict]:
    url = app_settings.drive_url()
    if not url:
        log("RESULT: 구글 드라이브 폴더 주소가 없습니다.")
        sys.exit(3)
    hub.ensure_chrome()
    DL_DIR.mkdir(exist_ok=True)
    files: list[dict] = []
    with sync_playwright() as p:
        b = p.chromium.connect_over_cdp(f"http://127.0.0.1:{hub.PORT}")
        ctx = b.contexts[0]
        page = next(
            (pg for pg in ctx.pages if not pg.is_closed() and "drive.google.com" in (pg.url or "")),
            None,
        )
        owned = page is None
        if page is None:
            page = ctx.new_page()
        try:
            _open_folder(page, url)
            if "accounts.google.com" in (page.url or "") or "signin" in (page.url or "").lower():
                log("RESULT: 구글 로그인이 필요합니다. Chrome에서 드라이브에 로그인하세요.")
                sys.exit(2)
            log("드라이브 폴더:", (page.title() or "")[:60])
            raw = page.evaluate(LIST_JS) or []
            files = _filter_xlsx(raw)
            if not files:
                # 이름만 xlsx인 항목도 후보로
                files = [
                    it for it in raw
                    if str(it.get("name") or "").lower().endswith(".xlsx")
                    and "shipmentsupload.xlsx" not in str(it.get("name") or "").lower()
                ]
            log(f"xlsx {len(files)}개")
            for it in files:
                log(" -", it.get("name"))
            if not files:
                try:
                    page.screenshot(path=str(SNAP))
                    log(f"목록이 비었습니다. 스크린샷: {SNAP}")
                except Exception:
                    pass
            _save_cache(url, files)
            if download and files:
                _download(ctx, page, files)
        finally:
            if owned:
                try:
                    if not page.is_closed():
                        page.close()
                except Exception:
                    pass
    log("RESULT: 드라이브 목록 완료")
    return files


def _download(ctx, page, files: list[dict]) -> list[Path]:
    saved: list[Path] = []
    for old in DL_DIR.glob("*.xlsx"):
        try:
            old.unlink()
        except Exception:
            pass
    for it in files:
        fid = it.get("id") or ""
        name = it.get("name") or f"{fid}.xlsx"
        if not name.lower().endswith(".xlsx"):
            name += ".xlsx"
        dest = DL_DIR / name
        url = f"https://drive.google.com/uc?export=download&id={fid}&confirm=t"
        try:
            resp = ctx.request.get(url, timeout=120000)
            body = resp.body()
            if not body or body[:15].lower().startswith(b"<!doctype") or b"<html" in body[:80].lower():
                log(f"직접 다운로드 실패, UI로 재시도: {name}")
                path = _download_ui(page, fid, dest)
            else:
                dest.write_bytes(body)
                path = dest
            if path and path.exists() and path.stat().st_size > 100:
                saved.append(path)
                log(f"받음: {path.name} ({path.stat().st_size} bytes)")
            else:
                log(f"다운로드 실패: {name}")
        except Exception as e:
            log(f"다운로드 오류 {name}: {e}")
    return saved


def _download_ui(page, fid: str, dest: Path) -> Path | None:
    try:
        row = page.locator(f'[data-id="{fid}"]').first
        if not row.count():
            return None
        row.click(button="right", timeout=4000)
        time.sleep(0.5)
        with page.expect_download(timeout=60000) as dl:
            clicked = False
            for t in ("다운로드", "Download"):
                loc = page.get_by_text(t, exact=True)
                if loc.count():
                    loc.first.click()
                    clicked = True
                    break
            if not clicked:
                return None
        dl.value.save_as(str(dest))
        return dest
    except Exception as e:
        log(f"UI 다운로드 실패: {e}")
        return None


def upload_local_files(ctx, paths: list[Path], folder_url: str | None = None) -> int:
    """로그인된 Chrome으로 지정한 드라이브 폴더에 파일을 올린다."""
    url = (folder_url or "").strip() or app_settings.drive_url()
    if not url:
        log("드라이브 주소가 없어 업로드를 건너뜁니다.")
        return 0
    files = [p for p in paths if p.exists()]
    if not files:
        return 0
    page = ctx.new_page()
    ok = 0
    try:
        page.goto(url, wait_until="domcontentloaded")
        time.sleep(5)
        if "accounts.google.com" in (page.url or "") or "signin" in (page.url or "").lower():
            log("RESULT: 구글 로그인이 필요합니다.")
            return 0
        log("드라이브 업로드 폴더:", (page.title() or "")[:50])
        for path in files:
            if _upload_one(page, path):
                ok += 1
        log(f"드라이브 업로드 {ok}/{len(files)}개")
        return ok
    finally:
        try:
            page.close()
        except Exception:
            pass


def _upload_one(page, path: Path) -> bool:
    new_btn = None
    for t in ("새로 만들기", "신규", "New"):
        try:
            btn = page.get_by_role("button", name=t)
            if btn.count():
                new_btn = btn.first
                break
        except Exception:
            pass
        try:
            el = page.get_by_text(t, exact=True)
            if el.count():
                new_btn = el.first
                break
        except Exception:
            pass
    if new_btn is None:
        log("새로 만들기 버튼을 못 찾았습니다.")
        return False
    new_btn.click()
    time.sleep(1.5)
    pat = re.compile("파일 업로드|File upload")
    target = None
    try:
        mi = page.get_by_role("menuitem", name=pat)
        if mi.count():
            target = mi.first
    except Exception:
        pass
    if target is None:
        try:
            tx = page.get_by_text(pat)
            if tx.count():
                target = tx.first
        except Exception:
            pass
    if target is None:
        log(f"파일 업로드 메뉴 없음: {path.name}")
        return False
    try:
        with page.expect_file_chooser(timeout=10000) as fc:
            target.click()
        fc.value.set_files(str(path))
    except Exception as e:
        log(f"파일 지정 실패 {path.name}: {e}")
        return False
    log(f"드라이브에 지정: {path.name}")
    for _ in range(45):
        time.sleep(2)
        try:
            txt = page.evaluate("() => document.body.innerText") or ""
        except Exception:
            txt = ""
        if "업로드 완료" in txt or "업로드됨" in txt or "upload complete" in txt.lower():
            log(f"드라이브 업로드 완료: {path.name}")
            return True
    log(f"완료 신호는 못 봤지만 지정됨: {path.name}")
    return True


def local_excels() -> list[Path]:
    if not DL_DIR.exists():
        return []
    files = [p for p in DL_DIR.glob("*.xlsx") if app_settings.is_ship_excel(p.name)]
    files.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return files


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else "list"
    list_and_download(download=(cmd == "download"))
