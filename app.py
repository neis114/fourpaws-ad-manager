"""네이버 검색광고(파워링크) 일괄관리 웹 대시보드.

- 톤앤매너: BAT 브랜드 아이덴티티 기반 High-end Black & White (Monochrome & Minimalist)
- 지원 기능: 키워드·소재·확장소재 6종 등록/수정/삭제/입찰가 변경, 미리보기, 실행, 현황 스냅샷
"""

import datetime as dt
import os
import sys

import pandas as pd
import streamlit as st

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
if CURRENT_DIR not in sys.path:
    sys.path.insert(0, CURRENT_DIR)

from ads_core import config, runner, specs  # noqa: E402
from ads_core.excel_io import read_jobs  # noqa: E402


# ---------------------------------------------------------
# 1. 페이지 설정 및 BAT Black & White 프리미엄 CSS
# ---------------------------------------------------------
st.set_page_config(
    page_title="BAT Ad Ops | Fourpaws",
    page_icon="■",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    @import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard/dist/web/static/pretendard.css');
    @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@700;900&display=swap');

    /* 폰트 및 안티앨리어싱: 아이콘 글리프는 건드리지 않고 텍스트만 적용 */
    .stApp, .stMarkdown, p, h1, h2, h3, h4, label, input, textarea {
        font-family: 'Pretendard', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
        -webkit-font-smoothing: antialiased;
        -moz-osx-font-smoothing: grayscale;
        letter-spacing: -0.015em;
    }

    /* 스트림릿 기본 아이콘 글리프 폰트 보호 (upload 글자 겹침 방지) */
    [data-testid="stIconMaterial"], .material-symbols-rounded, [class*="material-symbols"] {
        font-family: 'Material Symbols Rounded' !important;
        letter-spacing: normal !important;
    }

    /* 스트림릿 기본 헤더/푸터 숨김 */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    .stDeployButton {display: none;}

    /* 전체 배경: 세련된 퓨어 라이트 그레이 */
    .stApp {
        background-color: #FAFAFA;
        color: #000000;
    }

    /* 사이드바 */
    section[data-testid="stSidebar"] {
        background-color: #FFFFFF;
        border-right: 1px solid #E4E4E7;
        padding-top: 1.5rem;
    }

    /* BAT 타이포그래피 로고 컴포넌트 */
    .bat-brand-mark {
        font-family: 'Space Grotesk', 'Pretendard', sans-serif !important;
        font-weight: 900;
        line-height: 0.82;
        font-size: 28px;
        letter-spacing: -0.06em;
        color: #000000;
        display: inline-block;
        user-select: none;
    }

    /* 모노크롬 카드 컴포넌트 */
    .bat-card {
        background: #FFFFFF;
        border: 1px solid #E4E4E7;
        border-radius: 4px;
        padding: 20px 24px;
        margin-bottom: 20px;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.02);
    }

    .bat-card-title {
        font-size: 1.05rem;
        font-weight: 800;
        color: #000000;
        letter-spacing: -0.03em;
        margin-bottom: 6px;
    }

    .bat-card-desc {
        font-size: 0.86rem;
        color: #71717A;
        margin-bottom: 16px;
        line-height: 1.5;
    }

    /* 블랙 & 화이트 스탯 지표 */
    .bat-stat-box {
        background: #FFFFFF;
        border: 1px solid #000000;
        border-radius: 4px;
        padding: 16px 12px;
        text-align: center;
        transition: all 0.15s ease;
    }
    .bat-stat-box:hover {
        background: #000000;
        color: #FFFFFF !important;
    }
    .bat-stat-box:hover .stat-lbl,
    .bat-stat-box:hover .stat-num {
        color: #FFFFFF !important;
    }

    .stat-lbl {
        font-size: 0.78rem;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.04em;
        color: #71717A;
        margin-bottom: 4px;
    }
    .stat-num {
        font-size: 1.5rem;
        font-weight: 900;
        color: #000000;
        font-family: 'Space Grotesk', sans-serif !important;
    }

    /* 버튼 스타일 (BAT Solid Black) */
    div.stButton > button {
        border-radius: 4px;
        font-weight: 700;
        font-size: 0.88rem;
        padding: 0.6rem 1.2rem;
        transition: all 0.15s ease;
        border: 1px solid #000000;
        background-color: #FFFFFF;
        color: #000000;
    }
    div.stButton > button:hover {
        background-color: #000000;
        color: #FFFFFF;
        border-color: #000000;
    }

    /* 메인 블랙 버튼 */
    div.stButton > button[kind="primary"] {
        background: #000000 !important;
        color: #FFFFFF !important;
        border: 1px solid #000000 !important;
    }
    div.stButton > button[kind="primary"]:hover {
        background: #27272A !important;
        border-color: #27272A !important;
    }

    /* 탭 헤더 */
    .stTabs [data-baseweb="tab-list"] {
        gap: 0px;
        background-color: transparent;
        border-bottom: 1px solid #E4E4E7;
    }
    .stTabs [data-baseweb="tab"] {
        border-radius: 0;
        padding: 10px 20px;
        font-weight: 700;
        font-size: 0.92rem;
        color: #71717A;
        border: none;
        background: transparent;
    }
    .stTabs [aria-selected="true"] {
        color: #000000 !important;
        border-bottom: 2px solid #000000 !important;
        background: transparent !important;
    }

    /* 엑셀 업로더 커스텀: 글자 겹침 방지 */
    [data-testid="stFileUploader"] {
        background: #FFFFFF;
        border: 1px dashed #A1A1AA;
        border-radius: 4px;
        padding: 14px;
    }
    [data-testid="stFileUploader"] section {
        padding: 10px 0;
    }
    [data-testid="stFileUploader"] button {
        border-radius: 4px;
        border: 1px solid #000000;
        font-weight: 600;
    }

    /* 로그인 상자 */
    .bat-login-container {
        max-width: 400px;
        margin: 80px auto;
        background: #FFFFFF;
        padding: 40px 32px;
        border: 1px solid #000000;
        border-radius: 4px;
        text-align: center;
    }
    .bat-login-title {
        font-size: 1.15rem;
        font-weight: 900;
        letter-spacing: -0.03em;
        margin-top: 18px;
        margin-bottom: 6px;
    }
    .bat-login-sub {
        font-size: 0.82rem;
        color: #71717A;
        margin-bottom: 26px;
    }
</style>
""", unsafe_allow_html=True)


# ---------------------------------------------------------
# 2. BAT 시그니처 로그인 모달
# ---------------------------------------------------------
def check_auth():
    expected_pw = None
    if hasattr(st, "secrets") and "APP_PASSWORD" in st.secrets:
        expected_pw = str(st.secrets["APP_PASSWORD"]).strip()

    if not expected_pw:
        return True

    if st.session_state.get("authenticated", False):
        return True

    cols = st.columns([1, 2, 1])
    with cols[1]:
        st.markdown("""
        <div class="bat-login-container">
            <div class="bat-brand-mark">B/<br>AT</div>
            <div class="bat-login-title">PERFORMANCE AD OPS</div>
            <div class="bat-login-sub">접근 권한 확인을 위해 시스템 비밀번호를 입력하십시오.</div>
        </div>
        """, unsafe_allow_html=True)

        with st.form("bat_login_form"):
            input_pw = st.text_input("PASSWORD", type="password", placeholder="Enter passcode", label_visibility="collapsed")
            submit = st.form_submit_button("AUTHENTICATE", type="primary", use_container_width=True)
            if submit:
                if input_pw == expected_pw:
                    st.session_state["authenticated"] = True
                    st.rerun()
                else:
                    st.error("인증 실패: 비밀번호가 올바르지 않습니다.")
    return False


if not check_auth():
    st.stop()


# ---------------------------------------------------------
# 3. 환경변수(API 키) 취합
# ---------------------------------------------------------
def get_env_dict():
    env = {}
    if hasattr(st, "secrets"):
        for k, v in st.secrets.items():
            if k.startswith("NAVER_"):
                env[k] = str(v).strip().strip('"').strip("'")
    if not env:
        try:
            env = config.load_env()
        except Exception:
            env = {}
    return env


ENV = get_env_dict()
AVAILABLE_ACCOUNTS = config.account_codes(ENV)


# ---------------------------------------------------------
# 4. 사이드바 (Minimalist Black & White)
# ---------------------------------------------------------
with st.sidebar:
    st.markdown("""
    <div style="display: flex; align-items: center; gap: 14px; margin-bottom: 24px;">
        <div class="bat-brand-mark" style="font-size: 32px;">B/<br>AT</div>
        <div>
            <div style="font-weight: 900; font-size: 1.05rem; letter-spacing: -0.04em;">AD AUTOMATION</div>
            <div style="font-size: 0.76rem; color: #71717A; font-weight: 600;">PETDOC / FOURPAWS</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # 1) 연결된 계정 블록
    if AVAILABLE_ACCOUNTS:
        badges_html = "".join(
            f'<span style="display:inline-block; border: 1px solid #000; color:#000; padding:2px 7px; border-radius:2px; font-size:0.75rem; font-weight:700; margin:2px;">{acc}</span>'
            for acc in AVAILABLE_ACCOUNTS
        )
    else:
        badges_html = '<span style="font-size: 0.8rem; color: #71717A;">연결된 계정 없음 (Secrets 설정 필요)</span>'

    st.markdown(f"""
    <div class="bat-card" style="padding: 14px 16px; margin-bottom: 14px;">
        <div style="font-size: 0.78rem; font-weight: 800; text-transform: uppercase; letter-spacing: 0.05em; color: #71717A; margin-bottom: 8px;">CONNECTED ACCOUNTS</div>
        <div>{badges_html}</div>
    </div>
    """, unsafe_allow_html=True)

    # 2) 템플릿 다운로드 블록
    st.markdown("""
    <div class="bat-card" style="padding: 14px 16px; margin-bottom: 14px;">
        <div style="font-size: 0.78rem; font-weight: 800; text-transform: uppercase; letter-spacing: 0.05em; color: #71717A; margin-bottom: 4px;">STANDARD TEMPLATE</div>
        <div style="font-size: 0.78rem; color: #71717A; margin-bottom: 10px;">키워드·소재·확장소재 6종 규격</div>
    </div>
    """, unsafe_allow_html=True)

    template_path = os.path.join(CURRENT_DIR, "템플릿", "광고_일괄관리_템플릿_확장소재포함.xlsx")
    if os.path.exists(template_path):
        with open(template_path, "rb") as f:
            st.download_button(
                label="양식 내려받기 (.xlsx)",
                data=f.read(),
                file_name="광고_일괄관리_템플릿_확장소재포함.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

    if st.session_state.get("authenticated", False) and hasattr(st, "secrets") and "APP_PASSWORD" in st.secrets:
        st.write("")
        if st.button("LOGOUT", use_container_width=True):
            st.session_state["authenticated"] = False
            st.rerun()


# ---------------------------------------------------------
# 5. 메인 대시보드 탭
# ---------------------------------------------------------
tab_manage, tab_snapshot = st.tabs([
    "01. 일괄 관리 (BATCH OPS)",
    "02. 현황 스냅샷 (SNAPSHOT)"
])


# ---------------------------------------------------------
# TAB 1: 광고 일괄관리
# ---------------------------------------------------------
with tab_manage:
    st.markdown("""
    <div class="bat-card">
        <div class="bat-card-title">작업 파일 업로드 및 계획 수립</div>
        <div class="bat-card-desc">
            표준 양식에 기재된 작업 엑셀을 업로드하고 <b>[실행 계획 시뮬레이션]</b>을 실행하십시오.<br>
            네이버 검색광고 API에 즉시 반영되지 않으며, 등록·수정·삭제 및 문법 오류 여부를 사전 검증합니다.
        </div>
    </div>
    """, unsafe_allow_html=True)

    uploaded_file = st.file_uploader(
        "작업 파일",
        type=["xlsx"],
        label_visibility="collapsed"
    )

    if uploaded_file is not None:
        os.makedirs(config.INPUT_DIR, exist_ok=True)
        os.makedirs(config.OUTPUT_DIR, exist_ok=True)
        save_path = os.path.join(config.INPUT_DIR, uploaded_file.name)
        with open(save_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        col1, col2 = st.columns([1, 4])
        with col1:
            run_preview = st.button("실행 계획 시뮬레이션", type="primary", use_container_width=True)

        if run_preview:
            log_messages = []
            def web_log(msg):
                log_messages.append(str(msg))

            with st.spinner("구조 검증 및 네이버 계획 수립 중..."):
                try:
                    plans = runner.preview(save_path, env=ENV, log=web_log)
                    st.session_state["last_plans"] = plans
                    st.session_state["last_file"] = save_path
                    st.session_state["last_file_name"] = uploaded_file.name
                    st.session_state["preview_logs"] = log_messages
                except Exception as e:
                    st.error(f"시뮬레이션 실패: {e}")
                    st.session_state.pop("last_plans", None)

    # 시뮬레이션 결과 표시
    if "last_plans" in st.session_state and st.session_state.get("last_file_name") == getattr(uploaded_file, "name", None):
        plans = st.session_state["last_plans"]
        save_path = st.session_state["last_file"]

        ok_plans = [p for p in plans if not p.error]
        c_reg = sum(1 for p in ok_plans if p.action == "등록")
        c_mod = sum(1 for p in ok_plans if p.action == "수정")
        c_del = sum(1 for p in ok_plans if p.action == "삭제")
        c_err = len(plans) - len(ok_plans)

        st.write("")
        m1, m2, m3, m4 = st.columns(4)
        with m1:
            st.markdown(f'<div class="bat-stat-box"><div class="stat-lbl">CREATE</div><div class="stat-num">{c_reg}</div></div>', unsafe_allow_html=True)
        with m2:
            st.markdown(f'<div class="bat-stat-box"><div class="stat-lbl">UPDATE</div><div class="stat-num">{c_mod}</div></div>', unsafe_allow_html=True)
        with m3:
            st.markdown(f'<div class="bat-stat-box"><div class="stat-lbl">DELETE</div><div class="stat-num">{c_del}</div></div>', unsafe_allow_html=True)
        with m4:
            st.markdown(f'<div class="bat-stat-box"><div class="stat-lbl">ERROR</div><div class="stat-num">{c_err}</div></div>', unsafe_allow_html=True)

        # 세부 테이블
        rows = []
        for p in plans:
            rows.append({
                "대상": p.kind,
                "행번호": p.row,
                "계정": p.account,
                "작업": p.action or "오류",
                "항목명": p.job.get("기존 값") or p.job.get("키워드") or p.job.get("제목") or "",
                "상세 계획 / 사유": p.preview()
            })
        df_plans = pd.DataFrame(rows)

        st.write("")
        st.dataframe(df_plans, use_container_width=True, hide_index=True)

        preview_file = os.path.join(config.OUTPUT_DIR, f"{runner._stem(save_path)}_미리보기.xlsx")
        if os.path.exists(preview_file):
            with open(preview_file, "rb") as f:
                st.download_button(
                    label="시뮬레이션 검증 엑셀 내려받기",
                    data=f.read(),
                    file_name=os.path.basename(preview_file),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        # 네이버 실제 반영 구역
        st.write("")
        st.markdown("""
        <div class="bat-card" style="border: 1px solid #000000;">
            <div class="bat-card-title">네이버 검색광고 실제 운영 반영 (EXECUTION)</div>
            <div class="bat-card-desc">
                검증된 계획을 네이버 검색광고 서버에 직접 호출하여 실시간 반영합니다.<br>
                반영 전 스냅샷 백업(JSON)이 자동 생성되며, 기존 데이터 사본 파일에 최종 결과가 기록됩니다.
            </div>
        </div>
        """, unsafe_allow_html=True)

        confirm_check = st.checkbox("위 시뮬레이션 내역을 최종 확인하였으며, 네이버 광고 계정에 실제 변경을 승인합니다.")

        col_exec, _ = st.columns([1, 3])
        with col_exec:
            btn_apply = st.button("네이버 실제 반영 실행", type="primary", disabled=not confirm_check, use_container_width=True)

        if btn_apply:
            apply_logs = []
            def apply_web_log(msg):
                apply_logs.append(str(msg))

            with st.status("네이버 검색광고 API 반영 진행 중...", expanded=True) as status_box:
                try:
                    res = runner.apply(save_path, env=ENV, log=apply_web_log, confirm=lambda _: True)
                    status_box.update(label="반영 완료", state="complete")
                    st.success(f"처리 완료: 성공 {res.get('done', 0)}행 / 실패 {res.get('failed', 0)}행 / 사전 오류 {res.get('errors', 0)}행")

                    latest_results = sorted(
                        [os.path.join(config.OUTPUT_DIR, f) for f in os.listdir(config.OUTPUT_DIR) if f.startswith(f"{runner._stem(save_path)}_결과_")],
                        key=os.path.getmtime,
                        reverse=True
                    )
                    if latest_results and os.path.exists(latest_results[0]):
                        with open(latest_results[0], "rb") as f:
                            st.download_button(
                                label="최종 실행 결과 엑셀 내려받기 (.xlsx)",
                                data=f.read(),
                                file_name=os.path.basename(latest_results[0]),
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                type="primary"
                            )
                except Exception as e:
                    status_box.update(label="반영 중단", state="error")
                    st.error(f"실행 중 치명적 오류 발생: {e}")

                with st.expander("실행 상세 로그 확인", expanded=True):
                    st.code("\n".join(apply_logs))


# ---------------------------------------------------------
# TAB 2: 광고 현황 스냅샷 (Snapshot)
# ---------------------------------------------------------
with tab_snapshot:
    st.markdown("""
    <div class="bat-card">
        <div class="bat-card-title">네이버 파워링크 실시간 현황 스냅샷</div>
        <div class="bat-card-desc">
            광고 계정의 전체 키워드, 소재, 확장소재 6종 및 고유 시스템 ID를 실시간으로 조회하여 엑셀 파일로 추출합니다.<br>
            추출된 ID는 작업 템플릿 작성 시 타겟 지정용으로 활용할 수 있습니다.
        </div>
    </div>
    """, unsafe_allow_html=True)

    if not AVAILABLE_ACCOUNTS:
        st.warning("등록된 광고 계정 정보가 없습니다. Streamlit Secrets 설정을 확인하십시오.")
    else:
        selected_accounts = st.multiselect(
            "대상 광고 계정 선택",
            options=AVAILABLE_ACCOUNTS,
            default=AVAILABLE_ACCOUNTS[:2] if len(AVAILABLE_ACCOUNTS) >= 2 else AVAILABLE_ACCOUNTS
        )

        col_snap, _ = st.columns([1, 3])
        with col_snap:
            btn_snap = st.button("실시간 현황 엑셀 생성", type="primary", use_container_width=True)

        if btn_snap and selected_accounts:
            snap_logs = []
            def snap_web_log(msg):
                snap_logs.append(str(msg))

            with st.status("네이버 계정 전체 데이터 수집 및 엑셀 작성 중...", expanded=True) as snap_status:
                try:
                    out_path = runner.snapshot(selected_accounts, env=ENV, log=snap_web_log)
                    snap_status.update(label="스냅샷 작성 완료", state="complete")
                    if out_path and os.path.exists(out_path):
                        with open(out_path, "rb") as f:
                            st.download_button(
                                label="수집된 현황 엑셀 내려받기 (.xlsx)",
                                data=f.read(),
                                file_name=os.path.basename(out_path),
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                type="primary"
                            )
                except Exception as e:
                    snap_status.update(label="수집 실패", state="error")
                    st.error(f"스냅샷 수집 중 오류: {e}")

                with st.expander("데이터 수집 로그", expanded=False):
                    st.code("\n".join(snap_logs))
