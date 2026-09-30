# -*- coding: utf-8 -*-
"""대시보드 설정. 시트 주소, 쉽먼트/INBOX 폴더, 입고예정일을 저장한다.

값은 settings.json 에 저장된다. 이 파일은 사람마다 다르므로 업데이트 때 덮어쓰지 않는다.
"""
from __future__ import annotations

import json
import re
import sys
from pathlib import Path

BASE = Path(__file__).parent
SETTINGS_PATH = BASE / "settings.json"
DRIVE_CACHE = BASE / "drive_ship_cache.json"
DEFAULTS = {
    # 로켓발주·쉽먼트 계획 메뉴(Apps Script)가 들어 있는 본인 구글시트
    "order_sheet_url": "",
    # 쉽먼트 엑셀 폴더: 구글 드라이브 폴더 주소 또는 PC 폴더 경로
    "ship_folder": "",
    # 발주서 만들기에서 허브 상품목록을 올릴 INBOX 드라이브 폴더
    "inbox_folder": "",
    "inbound_date": "",
}

SHIP_NAME_KEYS = ("shipment_upload", "shipmentsupload", "쉽먼트", "shipment")
DRIVE_FOLDER_RE = re.compile(
    r"drive\.google\.com/drive/(?:u/\d+/)?folders/([a-zA-Z0-9_-]+)",
    re.I,
)
DRIVE_ID_RE = re.compile(r"[?&]id=([a-zA-Z0-9_-]+)")
SHEET_ID_RE = re.compile(r"/spreadsheets/d/([a-zA-Z0-9_-]+)")


def load() -> dict:
    data = dict(DEFAULTS)
    if SETTINGS_PATH.exists():
        try:
            data.update(json.loads(SETTINGS_PATH.read_text(encoding="utf-8")))
        except Exception:
            pass
    return data


def save(patch: dict) -> dict:
    data = load()
    data.update(patch)
    SETTINGS_PATH.write_text(
        json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return data


def inbound_date() -> str:
    return (load().get("inbound_date") or "").strip()


# ── 구글시트 ──────────────────────────────────────────────
def order_sheet_url() -> str:
    return (load().get("order_sheet_url") or "").strip()


def order_sheet_id() -> str:
    m = SHEET_ID_RE.search(order_sheet_url())
    return m.group(1) if m else ""


def require_order_sheet() -> str:
    """시트 주소가 없으면 안내 후 종료한다."""
    url = order_sheet_url()
    if not order_sheet_id():
        print("RESULT: 설정 탭에서 '발주 시트 주소'를 먼저 저장해 주세요.", flush=True)
        sys.exit(2)
    return url


# ── 드라이브 / 폴더 ───────────────────────────────────────
def ship_folder_raw() -> str:
    return (load().get("ship_folder") or "").strip()


def inbox_folder_raw() -> str:
    return (load().get("inbox_folder") or "").strip()


def inbox_drive_url() -> str:
    raw = inbox_folder_raw()
    if parse_drive_id(raw):
        return drive_url(raw)
    return raw


def parse_drive_id(raw: str) -> str | None:
    s = (raw or "").strip()
    if not s:
        return None
    m = DRIVE_FOLDER_RE.search(s)
    if m:
        return m.group(1)
    if "drive.google.com" in s.lower():
        m = DRIVE_ID_RE.search(s)
        if m:
            return m.group(1)
    if re.fullmatch(r"[a-zA-Z0-9_-]{25,}", s) and not Path(s).exists():
        return s
    return None


def is_drive(raw: str | None = None) -> bool:
    return parse_drive_id(raw if raw is not None else ship_folder_raw()) is not None


def drive_url(raw: str | None = None) -> str:
    fid = parse_drive_id(raw if raw is not None else ship_folder_raw())
    return f"https://drive.google.com/drive/folders/{fid}" if fid else ""


def ship_folder() -> Path:
    raw = ship_folder_raw()
    if is_drive(raw):
        return BASE / "_ship_dl"
    if not raw:
        return BASE / "쉽먼트"
    p = Path(raw)
    return p if p.is_absolute() else (BASE / p)


def is_ship_excel(name: str) -> bool:
    n = name.lower()
    if n.startswith("~$") or n == "shipmentsupload.xlsx":
        return False
    return any(k in n for k in SHIP_NAME_KEYS)


def list_ship_excels() -> list[Path]:
    folder = ship_folder()
    if not folder.exists() or not folder.is_dir():
        return []
    files = [
        p for p in folder.glob("*.xlsx")
        if is_ship_excel(p.name)
    ]
    files.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return files
