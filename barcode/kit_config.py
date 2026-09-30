# -*- coding: utf-8 -*-
"""바코드 앱 사용자 설정.

- 웹(Streamlit Cloud): 로그인한 사용자의 설정 (accounts.py 가 관리용 시트에서 읽어 세션에 넣는다)
- PC(비밀설정 없음):   ../settings.json

코드 어디에도 개인 시트 주소·업체 정보를 기본값으로 넣지 않는다.
"""
from __future__ import annotations

import json
from pathlib import Path

import streamlit as st

KIT_DIR = Path(__file__).resolve().parent.parent
SETTINGS_PATH = KIT_DIR / "settings.json"
LOCAL_KEY_PATH = KIT_DIR / "service_account.json"
SESSION_KEY = "kit_settings"


# ── 저장소 ────────────────────────────────────────────────
def is_web() -> bool:
    """관리용 시트가 비밀설정에 있으면 웹(로그인) 모드."""
    try:
        return bool(st.secrets.get("admin_sheet_url"))
    except Exception:  # secrets.toml 없음 → PC 모드
        return False


def _load_local() -> dict:
    if SETTINGS_PATH.exists():
        try:
            return json.loads(SETTINGS_PATH.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def load() -> dict:
    if SESSION_KEY not in st.session_state:
        st.session_state[SESSION_KEY] = {} if is_web() else _load_local()
    return st.session_state[SESSION_KEY]


def save(patch: dict) -> None:
    data = dict(load())
    data.update(patch)
    if is_web():
        import accounts
        accounts.save_settings(st.session_state["auth_user"]["id"], data)
    else:
        merged = _load_local()
        merged.update(patch)
        SETTINGS_PATH.write_text(json.dumps(merged, ensure_ascii=False, indent=2), encoding="utf-8")
    st.session_state[SESSION_KEY] = data


def get(key: str, default: str = "") -> str:
    v = load().get(key)
    return (str(v).strip() if v not in (None, "") else default)


# ── 시트 ──────────────────────────────────────────────────
def ops_sheet_url() -> str:
    """내 운영 시트 (필수). 아래 시트들을 비워 두면 이 시트를 쓴다."""
    return get("order_sheet_url")


def stock_sheet_url() -> str:
    return get("stock_sheet_url") or ops_sheet_url()


def stock_tab() -> str:
    return get("stock_tab", "등록상품정보")


def pick_ship_sheet_url() -> str:
    return get("pick_ship_sheet_url") or ops_sheet_url()


def pick_ship_tab() -> str:
    return get("pick_ship_tab", "출고확인")


def pick_dapae_sheet_url() -> str:
    """배대지 입고 시트 (선택). 비우면 쉽먼트 시트와 같은 시트."""
    return get("pick_dapae_sheet_url")


def pick_dapae_tab() -> str:
    return get("pick_dapae_tab", "배대지입고리스트")


def courier_addr_sheet_url() -> str:
    return get("courier_addr_sheet_url") or ops_sheet_url()


def courier_addr_tab() -> str:
    return get("courier_addr_tab", "주소리스트")


def service_account_path() -> Path:
    raw = get("service_account_path")
    return Path(raw) if raw else LOCAL_KEY_PATH


# ── 업체 · 라벨 ───────────────────────────────────────────
def company_name() -> str:
    return get("company_name")


LABEL_DEFAULTS = {
    # 소형 라벨
    "s_origin": "제조국 Made in China",
    "s_age": "본 제품은 14세 이상 사용가능합니다",
    # 대형 라벨
    "l_caution": "취급 상 주의 사항 : 화기에 주의 하세요.",
    "l_addr": "표시자 주소 및 전화번호 : (설정에서 입력)",
    "l_origin": "제조국 : Made in China",
    "l_age": "사용연령 : 만14세이상",
}
LABEL_COL_DEFAULTS = {"name": 4, "barcode": 12, "material": 13, "insert": 13, "startrow": 2}


def label_text(key: str) -> str:
    """라벨 고정 문구. 사용자가 저장한 값이 있으면 그 값."""
    return get(f"label_{key}", LABEL_DEFAULTS[key])


def label_col(prefix: str, name: str) -> int:
    """라벨 열 번호 (prefix: 's' 소형 / 'l' 대형)."""
    try:
        return int(load().get(f"label_{prefix}_{name}") or LABEL_COL_DEFAULTS[name])
    except (TypeError, ValueError):
        return LABEL_COL_DEFAULTS[name]


def label_address() -> str:
    return label_text("l_addr")


# ── 택배 양식 ─────────────────────────────────────────────
# 엑셀 한 줄 = 박스(송장) 1개. 각 칸에 넣을 수 있는 항목:
COURIER_FIELDS = {
    "보내는분": "주소리스트 A열 (비면 설정의 보내는 사람)",
    "보내는전화1": "주소리스트 B열 (비면 설정의 보내는 전화)",
    "보내는전화2": "주소리스트 C열",
    "보내는주소": "주소리스트 D열 (비면 설정의 보내는 주소)",
    "받는분": "물류센터-송장끝6자리 (수량개)",
    "받는전화1": "주소리스트 F열",
    "받는전화2": "주소리스트 G열",
    "받는우편": "주소리스트 H열",
    "받는주소": "주소리스트 I열",
    "물류센터": "출고확인의 물류센터",
    "송장번호": "쉽먼트 운송장번호",
    "수량": "박스 안 수량 합계",
    "빈칸": "비워 둠",
}

# 한진택배 대량 접수 양식 (기본값)
COURIER_LAYOUT_DEFAULT = [
    ["보낼분", "보내는분"],
    ["전화번호1", "보내는전화1"],
    ["전화번호2", "보내는전화2"],
    ["주소", "보내는주소"],
    ["받는분(박스명)", "받는분"],
    ["받는전화1", "받는전화1"],
    ["받는전화2", "받는전화2"],
    ["받는우편", "받는우편"],
    ["받는주소", "받는주소"],
]


def validate_courier_layout(rows) -> tuple[list[list[str]], list[str]]:
    cols, errors = [], []
    for i, row in enumerate(rows or [], start=1):
        row = list(row or []) + ["", ""]
        head = str(row[0] or "").strip()
        field = str(row[1] or "").strip()
        if not head and not field:
            continue
        if not head:
            errors.append(f"{i}번째 열의 머리글이 비었습니다")
        elif field not in COURIER_FIELDS:
            errors.append(f"{i}번째 열 '{head}' 의 항목을 고르세요")
        else:
            cols.append([head, field])
    if not cols and not errors:
        errors.append("열이 하나도 없습니다")
    return cols, errors


def courier_name() -> str:
    return get("courier_name", "한진택배")


def courier_layout() -> list[tuple[str, str]]:
    cols, errors = validate_courier_layout(load().get("courier_layout") or COURIER_LAYOUT_DEFAULT)
    if errors:
        cols = COURIER_LAYOUT_DEFAULT
    return [(h, f) for h, f in cols]


def sender_defaults() -> dict:
    return {
        "보내는분": get("sender_name"),
        "보내는전화1": get("sender_phone"),
        "보내는주소": get("sender_address"),
    }


# ── 상단 탭 순서 ───────────────────────────────────────────
def tab_order(all_keys: list[str]) -> list[str]:
    """저장된 순서대로. 새로 생긴 탭은 뒤에 붙인다."""
    saved = [k for k in (load().get("tab_order") or []) if k in all_keys]
    return saved + [k for k in all_keys if k not in saved]
