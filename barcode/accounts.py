# -*- coding: utf-8 -*-
"""웹 모드 계정 · 사용자 설정 저장소 (관리용 구글시트).

비밀설정(.streamlit/secrets.toml 또는 Streamlit Cloud Secrets)에 필요한 값:
    admin_sheet_url = "https://docs.google.com/spreadsheets/d/..."   # 관리용 시트
    [admin]
    id = "admin"
    password = "..."                                                 # 관리자 로그인
    gcp_service_account_json = '''{ ...credentials.json 내용 그대로... }'''  # 서비스 계정 키

관리용 시트의 탭 (없으면 자동으로 만든다):
    계정: 아이디 | 이름 | 비밀번호 | 권한 | 사용 | 만든날 | 마지막로그인
    설정: 아이디 | 설정 | 수정일
비밀번호는 PBKDF2 해시로만 저장한다.
"""
from __future__ import annotations

import datetime as dt
import hashlib
import hmac
import json
import re
import secrets as _secrets

import streamlit as st

USERS_TAB = "계정"
SETTINGS_TAB = "설정"
USERS_HEAD = ["아이디", "이름", "비밀번호", "권한", "사용", "만든날", "마지막로그인"]
SETTINGS_HEAD = ["아이디", "설정", "수정일"]
ID_RE = re.compile(r"^[a-zA-Z0-9_.-]{3,30}$")
ITERATIONS = 200_000


# ── 비밀번호 ──────────────────────────────────────────────
def hash_password(pw: str) -> str:
    salt = _secrets.token_hex(16)
    digest = hashlib.pbkdf2_hmac("sha256", pw.encode(), bytes.fromhex(salt), ITERATIONS).hex()
    return f"pbkdf2${ITERATIONS}${salt}${digest}"


def check_password(pw: str, stored: str) -> bool:
    try:
        _, n, salt, digest = stored.split("$")
        calc = hashlib.pbkdf2_hmac("sha256", pw.encode(), bytes.fromhex(salt), int(n)).hex()
        return hmac.compare_digest(calc, digest)
    except Exception:
        return False


def _now() -> str:
    return (dt.datetime.utcnow() + dt.timedelta(hours=9)).strftime("%Y-%m-%d %H:%M")


# ── 시트 ──────────────────────────────────────────────────
@st.cache_resource(ttl=3600)
def _client():
    import gspread
    from google.oauth2.service_account import Credentials

    creds = Credentials.from_service_account_info(
        service_account_info(),
        scopes=["https://www.googleapis.com/auth/spreadsheets",
                "https://www.googleapis.com/auth/drive"],
    )
    client = gspread.authorize(creds, http_client=gspread.BackOffHTTPClient)
    client.set_timeout((10, 60))
    return client


def service_account_info() -> dict:
    """Secrets 의 서비스 계정 키.

    - gcp_service_account_json = '''(credentials.json 내용 그대로)'''   ← 붙여 넣기만 하면 됨
    - [gcp_service_account] 표 형식도 그대로 지원
    """
    raw = st.secrets.get("gcp_service_account_json")
    if raw:
        return json.loads(raw) if isinstance(raw, str) else dict(raw)
    return dict(st.secrets["gcp_service_account"])


def has_service_account() -> bool:
    try:
        return bool(st.secrets.get("gcp_service_account_json") or "gcp_service_account" in st.secrets)
    except Exception:
        return False


def service_email() -> str:
    try:
        return service_account_info().get("client_email", "")
    except Exception:
        return ""


def _ws(tab: str, head: list[str]):
    book = _client().open_by_url(st.secrets["admin_sheet_url"])
    try:
        ws = book.worksheet(tab)
    except Exception:
        ws = book.add_worksheet(title=tab, rows=200, cols=len(head))
        ws.update([head], "A1")
    return ws


def _rows(tab: str, head: list[str]) -> tuple[object, list[list[str]]]:
    ws = _ws(tab, head)
    values = ws.get_all_values()
    return ws, values[1:] if values else []


# ── 계정 ──────────────────────────────────────────────────
def _admin_login(uid: str, pw: str) -> dict | None:
    try:
        admin = st.secrets["admin"]
    except Exception:
        return None
    if uid == admin.get("id") and admin.get("password") and hmac.compare_digest(pw, admin["password"]):
        return {"id": uid, "name": "관리자", "role": "admin"}
    return None


def login(uid: str, pw: str) -> tuple[dict | None, str]:
    uid = (uid or "").strip()
    if not uid or not pw:
        return None, "아이디와 비밀번호를 입력하세요."
    user = _admin_login(uid, pw)
    if user:
        return user, ""
    ws, rows = _rows(USERS_TAB, USERS_HEAD)
    for i, r in enumerate(rows, start=2):
        r = r + [""] * (len(USERS_HEAD) - len(r))
        if r[0] != uid:
            continue
        if not check_password(pw, r[2]):
            break
        if r[4].strip().upper() == "N":
            return None, "사용이 중지된 계정입니다. 강사에게 문의하세요."
        try:
            ws.update_cell(i, 7, _now())
        except Exception:
            pass
        return {"id": uid, "name": r[1] or uid, "role": r[3] or "user"}, ""
    return None, "아이디 또는 비밀번호가 틀렸습니다."


def list_users() -> list[dict]:
    _, rows = _rows(USERS_TAB, USERS_HEAD)
    out = []
    for r in rows:
        r = r + [""] * (len(USERS_HEAD) - len(r))
        if r[0]:
            out.append({"아이디": r[0], "이름": r[1], "권한": r[3] or "user",
                        "사용": r[4] or "Y", "만든날": r[5], "마지막로그인": r[6]})
    return out


def create_user(uid: str, name: str, pw: str) -> str:
    uid = (uid or "").strip()
    if not ID_RE.match(uid):
        return "아이디는 영문·숫자·_.- 3~30자입니다."
    if len(pw or "") < 6:
        return "비밀번호는 6자 이상입니다."
    ws, rows = _rows(USERS_TAB, USERS_HEAD)
    if any(r and r[0] == uid for r in rows):
        return "이미 있는 아이디입니다."
    ws.append_row([uid, name.strip(), hash_password(pw), "user", "Y", _now(), ""],
                  value_input_option="RAW")
    return ""


def _find_row(ws_rows, uid: str) -> int | None:
    for i, r in enumerate(ws_rows, start=2):
        if r and r[0] == uid:
            return i
    return None


def set_password(uid: str, pw: str) -> str:
    if len(pw or "") < 6:
        return "비밀번호는 6자 이상입니다."
    ws, rows = _rows(USERS_TAB, USERS_HEAD)
    row = _find_row(rows, uid)
    if row is None:
        return "계정을 찾지 못했습니다."
    ws.update_cell(row, 3, hash_password(pw))
    return ""


def set_active(uid: str, active: bool) -> str:
    ws, rows = _rows(USERS_TAB, USERS_HEAD)
    row = _find_row(rows, uid)
    if row is None:
        return "계정을 찾지 못했습니다."
    ws.update_cell(row, 5, "Y" if active else "N")
    return ""


# ── 사용자 설정 ───────────────────────────────────────────
def load_settings(uid: str) -> dict:
    _, rows = _rows(SETTINGS_TAB, SETTINGS_HEAD)
    for r in rows:
        if r and r[0] == uid and len(r) > 1 and r[1]:
            try:
                return json.loads(r[1])
            except Exception:
                return {}
    return {}


def save_settings(uid: str, data: dict) -> None:
    ws, rows = _rows(SETTINGS_TAB, SETTINGS_HEAD)
    blob = json.dumps(data, ensure_ascii=False)
    row = _find_row(rows, uid)
    if row is None:
        ws.append_row([uid, blob, _now()], value_input_option="RAW")
    else:
        ws.update([[blob, _now()]], f"B{row}:C{row}", value_input_option="RAW")
