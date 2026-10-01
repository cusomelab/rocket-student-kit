# -*- coding: utf-8 -*-
"""로그인 · 내 설정 · 관리자 화면."""
from __future__ import annotations

import re

import streamlit as st

import kit_config

SHEET_RE = re.compile(r"/spreadsheets/d/([a-zA-Z0-9_-]+)")


# ── 로그인 ────────────────────────────────────────────────
def require_login() -> None:
    """웹 모드에서 로그인하지 않았으면 로그인 화면을 보여주고 멈춘다."""
    if not kit_config.is_web():
        return
    problem = kit_config.secrets_problem()
    if problem:
        st.title("🚀 로켓배송 운영 관리")
        st.error("관리자 설정(Secrets) 오류로 앱을 열 수 없습니다. 강사에게 문의하세요.")
        st.caption(f"원인: {problem}")
        st.stop()
    if st.session_state.get("auth_user"):
        _sidebar_user()
        return

    import accounts

    st.title("🚀 로켓배송 운영 관리")
    st.caption("강사에게 받은 아이디와 비밀번호로 로그인하세요.")
    with st.form("login_form"):
        uid = st.text_input("아이디", key="login_id")
        pw = st.text_input("비밀번호", type="password", key="login_pw")
        ok = st.form_submit_button("로그인", type="primary", use_container_width=True)
    if ok:
        try:
            user, err = accounts.login(uid, pw)
        except Exception as e:
            user, err = None, f"로그인 서버 연결 실패: {e}"
        if user:
            try:
                settings = accounts.load_settings(user["id"])
            except Exception as e:
                st.error(f"내 설정을 불러오지 못했습니다. 잠시 후 다시 시도하세요. ({e})")
                st.stop()
            st.session_state["auth_user"] = user
            st.session_state[kit_config.SESSION_KEY] = settings
            st.rerun()
        else:
            st.error(err)
    st.stop()


def _sidebar_user() -> None:
    user = st.session_state["auth_user"]
    with st.sidebar:
        st.markdown(f"**{user['name']}** 님 (`{user['id']}`)")
        if st.button("로그아웃", key="logout_btn"):
            for k in list(st.session_state.keys()):
                del st.session_state[k]
            st.rerun()
        if not kit_config.ops_sheet_url():
            st.warning("⚙️ 설정 탭에서 내 운영 시트 주소를 먼저 저장하세요.")
        if user["role"] != "admin":
            with st.expander("비밀번호 변경"):
                old = st.text_input("지금 비밀번호", type="password", key="chpw_old")
                new = st.text_input("새 비밀번호 (6자 이상)", type="password", key="chpw_new")
                if st.button("변경", key="chpw_btn"):
                    import accounts
                    ok, _ = accounts.login(user["id"], old)
                    if not ok:
                        st.error("지금 비밀번호가 틀렸습니다.")
                    else:
                        err = accounts.set_password(user["id"], new)
                        if err:
                            st.error(err)
                        else:
                            st.success("변경했습니다.")


def is_admin() -> bool:
    user = st.session_state.get("auth_user") or {}
    return user.get("role") == "admin"


# ── 내 설정 ───────────────────────────────────────────────
def _sheet_input(label: str, key: str, placeholder: str = "비우면 운영 시트 사용") -> str:
    return st.text_input(label, value=kit_config.get(key), placeholder=placeholder, key=f"set_{key}")


def render_settings(tab_labels: dict[str, str], get_client) -> None:
    st.header("⚙️ 내 설정")
    st.caption("한 번 저장하면 다음 로그인부터 그대로 적용됩니다." if kit_config.is_web()
               else "이 PC의 settings.json 에 저장됩니다.")

    patch: dict = {}

    # ① 운영 시트
    st.subheader("① 내 운영 시트 (필수)")
    patch["order_sheet_url"] = st.text_input(
        "운영 시트 주소", value=kit_config.get("order_sheet_url"),
        placeholder="https://docs.google.com/spreadsheets/d/...", key="set_order_sheet_url")
    st.caption("강사 템플릿을 사본으로 만든 내 시트. 아래 시트들을 비워 두면 이 시트를 씁니다.")

    # ② 기능별 시트
    st.subheader("② 기능별 시트 · 탭")
    rows = [
        ("재고 확인", "stock_sheet_url", "stock_tab", "등록상품정보"),
        ("피킹 · 쉽먼트 시트", "pick_ship_sheet_url", "pick_ship_tab", "출고확인"),
        ("피킹 · 배대지 입고 (선택)", "pick_dapae_sheet_url", "pick_dapae_tab", "배대지입고리스트"),
        ("택배 주소리스트", "courier_addr_sheet_url", "courier_addr_tab", "주소리스트"),
    ]
    for label, url_key, tab_key, tab_default in rows:
        c1, c2 = st.columns([3, 1])
        with c1:
            ph = "비우면 쉽먼트 시트와 같은 시트" if url_key == "pick_dapae_sheet_url" else "비우면 운영 시트 사용"
            patch[url_key] = _sheet_input(label, url_key, ph)
        with c2:
            patch[tab_key] = st.text_input("탭 이름", value=kit_config.get(tab_key, tab_default), key=f"set_{tab_key}")

    email = ""
    if kit_config.is_web():
        import accounts
        email = accounts.service_email()
    if email:
        st.info(f"위 시트들을 모두 아래 이메일에 **편집자**로 공유하세요.\n\n`{email}`")

    # ③ 업체 · 라벨
    st.subheader("③ 업체 · 라벨 고정 문구")
    patch["company_name"] = st.text_input("업체명", value=kit_config.get("company_name"),
                                          placeholder="(주)업체명 — 공문·파일 이름에 사용", key="set_company_name")
    with st.expander("소형 라벨 고정 문구"):
        patch["label_s_origin"] = st.text_input("제조국", value=kit_config.label_text("s_origin"), key="set_label_s_origin")
        patch["label_s_age"] = st.text_input("사용연령", value=kit_config.label_text("s_age"), key="set_label_s_age")
    with st.expander("대형 라벨 고정 문구", expanded=True):
        patch["label_l_caution"] = st.text_area("취급주의", value=kit_config.label_text("l_caution"), height=80, key="set_label_l_caution")
        patch["label_l_addr"] = st.text_input("주소/전화", value=kit_config.label_text("l_addr"), key="set_label_l_addr")
        patch["label_l_origin"] = st.text_input("제조국", value=kit_config.label_text("l_origin"), key="set_label_l_origin")
        patch["label_l_age"] = st.text_input("사용연령", value=kit_config.label_text("l_age"), key="set_label_l_age")

    # ④ 택배 양식
    st.subheader("④ 택배 양식")
    st.caption("택배사 대량 접수 엑셀의 열 순서대로 머리글과 들어갈 항목을 정합니다. 한 줄 = 박스(송장) 1개.")
    patch["courier_name"] = st.text_input("택배사 이름", value=kit_config.courier_name(), key="set_courier_name")
    c1, c2, c3 = st.columns(3)
    with c1:
        patch["sender_name"] = st.text_input("보내는 사람", value=kit_config.get("sender_name"), key="set_sender_name",
                                             help="주소리스트 A열이 비었을 때 사용")
    with c2:
        patch["sender_phone"] = st.text_input("보내는 전화", value=kit_config.get("sender_phone"), key="set_sender_phone",
                                              help="주소리스트 B열이 비었을 때 사용")
    with c3:
        patch["sender_address"] = st.text_input("보내는 주소", value=kit_config.get("sender_address"), key="set_sender_address",
                                                help="주소리스트 D열이 비었을 때 사용")

    if st.button("↩️ 한진 기본 양식으로 되돌리기", key="courier_reset"):
        st.session_state["courier_editor_rows"] = [list(r) for r in kit_config.COURIER_LAYOUT_DEFAULT]
        st.session_state.pop("courier_editor", None)
    if "courier_editor_rows" not in st.session_state:
        st.session_state["courier_editor_rows"] = [list(r) for r in kit_config.courier_layout()]
    fields = list(kit_config.COURIER_FIELDS)
    edited = st.data_editor(
        [{"엑셀 머리글": h, "들어갈 항목": f} for h, f in st.session_state["courier_editor_rows"]],
        num_rows="dynamic",
        use_container_width=True,
        key="courier_editor",
        column_config={
            "엑셀 머리글": st.column_config.TextColumn(required=True),
            "들어갈 항목": st.column_config.SelectboxColumn(options=fields, required=True),
        },
    )
    with st.expander("항목 설명"):
        for f, desc in kit_config.COURIER_FIELDS.items():
            st.caption(f"**{f}** — {desc}")
    layout_rows = [[r.get("엑셀 머리글"), r.get("들어갈 항목")] for r in edited]
    layout, layout_errors = kit_config.validate_courier_layout(layout_rows)
    patch["courier_layout"] = layout

    # ⑤ 상단 탭 순서
    st.subheader("⑤ 상단 탭 순서")
    st.caption("위에서부터 순서대로 표시됩니다. 번호를 바꾸고 저장하세요.")
    order = kit_config.tab_order(list(tab_labels))
    order_rows = st.data_editor(
        [{"순서": i + 1, "탭": tab_labels[k], "_key": k} for i, k in enumerate(order)],
        use_container_width=True,
        hide_index=True,
        disabled=["탭"],
        column_config={"순서": st.column_config.NumberColumn(min_value=1, step=1), "_key": None},
        key="tab_order_editor",
    )
    patch["tab_order"] = [r["_key"] for r in sorted(order_rows, key=lambda r: (r.get("순서") or 999))]

    st.divider()
    errors = list(layout_errors)
    for k in ("order_sheet_url", "stock_sheet_url", "pick_ship_sheet_url",
              "pick_dapae_sheet_url", "courier_addr_sheet_url"):
        v = (patch.get(k) or "").strip()
        if v and not SHEET_RE.search(v):
            errors.append(f"시트 주소 형식이 아닙니다: {v[:60]}")
    if not (patch.get("order_sheet_url") or "").strip():
        errors.append("내 운영 시트 주소는 꼭 넣어야 합니다.")

    c1, c2 = st.columns([1, 1])
    with c1:
        if st.button("💾 설정 저장", type="primary", use_container_width=True, key="settings_save"):
            if errors:
                for e in errors:
                    st.error(e)
            else:
                clean = {k: (v.strip() if isinstance(v, str) else v) for k, v in patch.items()}
                try:
                    kit_config.save(clean)
                    st.session_state.pop("courier_editor_rows", None)
                    st.success("저장했습니다. 다음 로그인부터 이 값으로 동작합니다.")
                except Exception as e:
                    st.error(f"저장 실패: {e}")
    with c2:
        test = st.button("🔌 시트 연결 테스트 (저장된 설정 기준)", use_container_width=True, key="settings_test")
    if test:
        client = get_client()
        if client is None:
            st.error("구글 인증 정보를 읽지 못했습니다. (웹: 강사 문의 / PC: service_account.json 확인)")
        else:
            for name, ok, msg in connection_test(client, email):
                (st.success if ok else st.error)(f"{'✅' if ok else '❌'} {name} — {msg}")


def connection_test(client, email: str) -> list[tuple[str, bool, str]]:
    """저장된 설정 기준으로 시트·탭 접근을 확인한다."""
    checks = [
        ("재고 확인", kit_config.stock_sheet_url(), kit_config.stock_tab(), True),
        ("피킹 · 쉽먼트", kit_config.pick_ship_sheet_url(), kit_config.pick_ship_tab(), True),
        ("피킹 · 배대지", kit_config.pick_dapae_sheet_url() or kit_config.pick_ship_sheet_url(),
         kit_config.pick_dapae_tab(), False),
        ("택배 주소리스트", kit_config.courier_addr_sheet_url(), kit_config.courier_addr_tab(), True),
    ]
    out = []
    cache: dict[str, list[str] | str] = {}
    for name, url, tab, required in checks:
        if not url:
            out.append((name, not required, "시트 주소를 넣어 주세요." if required else "사용 안 함"))
            continue
        if url not in cache:
            try:
                cache[url] = [ws.title for ws in client.open_by_url(url).worksheets()]
            except Exception as e:
                cache[url] = sheet_error_hint(str(e), email)
        titles = cache[url]
        if isinstance(titles, str):
            out.append((name, False, titles))
        elif tab in titles:
            out.append((name, True, f"'{tab}' 탭 확인"))
        else:
            out.append((name, not required, f"'{tab}' 탭이 없습니다. 탭 이름을 확인해 주세요."))
    return out


def sheet_error_hint(msg: str, email: str) -> str:
    low = msg.lower()
    if "403" in msg and ("has not been used" in low or "disabled" in low):
        return "구글 클라우드에서 Google Sheets API · Drive API를 '사용'으로 켜 주세요. (강사 문의)"
    if "403" in msg or "permission" in low:
        return f"시트 오른쪽 위 [공유] → {email or '서비스 계정 이메일'} 을 '편집자'로 추가해 주세요."
    if "404" in msg or "not found" in low:
        return "시트 주소가 틀렸거나 삭제된 시트입니다. 주소를 다시 복사해 주세요."
    return f"연결 실패: {msg[:200]}"


# ── 관리자 ────────────────────────────────────────────────
def render_admin() -> None:
    import accounts

    st.header("👑 수강생 계정 관리")
    st.caption("계정은 관리용 시트의 '계정' 탭에 저장됩니다. 비밀번호는 암호화되어 저장됩니다.")

    with st.form("admin_create", clear_on_submit=True):
        st.subheader("새 계정 만들기")
        c1, c2, c3 = st.columns(3)
        with c1:
            uid = st.text_input("아이디 (영문·숫자)")
        with c2:
            name = st.text_input("이름")
        with c3:
            pw = st.text_input("처음 비밀번호 (6자 이상)", type="password")
        if st.form_submit_button("계정 만들기", type="primary"):
            err = accounts.create_user(uid, name, pw)
            st.error(err) if err else st.success(f"'{uid}' 계정을 만들었습니다.")

    st.subheader("계정 목록")
    try:
        users = accounts.list_users()
    except Exception as e:
        st.error(f"관리용 시트를 읽지 못했습니다: {e}")
        return
    if not users:
        st.info("아직 계정이 없습니다.")
        return
    st.dataframe(users, use_container_width=True, hide_index=True)

    st.subheader("비밀번호 초기화 · 사용 중지")
    ids = [u["아이디"] for u in users]
    c1, c2 = st.columns(2)
    with c1:
        target = st.selectbox("계정", ids, key="admin_target")
        new_pw = st.text_input("새 비밀번호", type="password", key="admin_new_pw")
        if st.button("비밀번호 바꾸기", key="admin_pw_btn"):
            err = accounts.set_password(target, new_pw)
            st.error(err) if err else st.success("바꿨습니다.")
    with c2:
        target2 = st.selectbox("계정 ", ids, key="admin_target2")
        cc1, cc2 = st.columns(2)
        with cc1:
            if st.button("사용 중지", key="admin_off"):
                err = accounts.set_active(target2, False)
                st.error(err) if err else st.success("중지했습니다.")
        with cc2:
            if st.button("다시 사용", key="admin_on"):
                err = accounts.set_active(target2, True)
                st.error(err) if err else st.success("다시 쓸 수 있습니다.")
