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

    html, body, [class*="css"], .stMarkdown, p, div, span, button, input {
        font-family: 'Pretendard', -apple-system, BlinkMacSystemFont, system-ui, sans-serif !important;
        letter-spacing: -0.02em;
    }

    /* 스트림릿 기본 요소 숨김 */
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
        padding-top: 1.8rem;
    }

    /* BAT 타이포그래피 로고 컴포넌트 */
    .bat-brand-mark {
        font-family: 'Space Grotesk', 'Pretendard', sans-serif !important;
        font-weight: 900;
        line-height: 0.82;
        font-size: 26px;
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
        padding: 24px;
        margin-bottom: 20px;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.02);
    }

    .bat-card-title {
        font-size: 1.05rem;
        font-weight: 800;
        color: #000000;
        letter-spacing: -0.03em;
        margin-bottom: 6px;
        display: flex;
        align-items: center;
        gap: 8px;
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

    /* 엑셀 업로더 커스텀 */
    [data-testid="stFileUploader"] {
        border: 1px dashed #A1A1AA;
        border-radius: 4px;
        padding: 12px;
        background: #FFFFFF;
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

        with st.form("bat_login_form", clear_on_submit=False):
            input_pw = st.text_input("PASSWORD", type="password", placeholder="비밀번호 입력", label_visibility="collapsed")
            submit = st.form_submit_button("인증 및 접속", use_container_width=True, type="primary")
            if submit:
                if input_pw == expected_pw:
                    st.session_state["authenticated"] = True
                    st.rerun()
                else:
                    st.error("비밀번호가 올바르지 않습니다.")
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

    st.markdown('<div class="bat-card" style="padding: 16px; margin-bottom: 12px;">', unsafe_allow_html=True)
    st.markdown('<div style="font-size: 0.78rem; font-weight: 800; text-transform: uppercase; letter-spacing: 0.05em; color: #71717A; margin-bottom: 8px;">CONNECTED ACCOUNTS</div>', unsafe_allow_html=True)
    if AVAILABLE_ACCOUNTS:
        badges = "".join(f'<span style="display:inline-block; border: 1px solid #000; color:#000; padding:2px 8px; border-radius:2px; font-size:0.75rem; font-weight:700; margin:2px;">{acc}</span>' for acc in AVAILABLE_ACCOUNTS)
        st.markdown(badges, unsafe_allow_html=True)
    else:
        st.markdown('<span style="font-size: 0.8rem; color: #000; font-weight: 700;">연결된 계정 없음 (Secrets 설정 필요)</span>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="bat-card" style="padding: 16px; margin-bottom: 12px;">', unsafe_allow_html=True)
    st.markdown('<div style="font-size: 0.78rem; font-weight: 800; text-transform: uppercase; letter-spacing: 0.05em; color: #71717A; margin-bottom: 6px;">STANDARD TEMPLATE</div>', unsafe_allow_html=True)
    st.markdown('<div style="font-size: 0.78rem; color: #71717A; margin-bottom: 12px;">키워드·소재·확장소재 6종 규격</div>', unsafe_allow_html=True)
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
    st.markdown('</div>', unsafe_allow_html=True)

    if st.session_state.get("authenticated", False) and hasattr(st, "secrets") and "APP_PASSWORD" in st.secrets:
        if st.button("LOGOUT", use_container_width=True):
            st.session_state["authenticated"] = False
            st.rerun()


# ---------------------------------------------------------
# 5. 메인 화면 탭
# ---------------------------------------------------------
tab_manage, tab_snapshot = st.tabs(["01. 일괄 관리 (BATCH OPS)", "02. 현황 스냅샷 (SNAPSHOT)"])


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

    # 시뮬레이션 결과
    if "last_plans" in st.session_state and st.session_state["last_plans"]:
        plans = st.session_state["last_plans"]
        stem_name = os.path.splitext(st.session_state.get("last_file_name", "작업"))[0]

        ok_plans = [p for p in plans if not p.error]
        c_reg = sum(1 for p in ok_plans if p.action == "등록")
        c_mod = sum(1 for p in ok_plans if p.action == "수정")
        c_del = sum(1 for p in ok_plans if p.action == "삭제")
        c_err = len(plans) - len(ok_plans)

        st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)

        # BAT 미니멀 모노크롬 스탯
        s1, s2, s3, s4 = st.columns(4)
        with s1:
            st.markdown(f'<div class="bat-stat-box"><div class="stat-lbl">CREATE</div><div class="stat-num">{c_reg}</div></div>', unsafe_allow_html=True)
        with s2:
            st.markdown(f'<div class="bat-stat-box"><div class="stat-lbl">UPDATE</div><div class="stat-num">{c_mod}</div></div>', unsafe_allow_html=True)
        with s3:
            st.markdown(f'<div class="bat-stat-box"><div class="stat-lbl">DELETE</div><div class="stat-num">{c_del}</div></div>', unsafe_allow_html=True)
        with s4:
            st.markdown(f'<div class="bat-stat-box" style="border-color: #71717A;"><div class="stat-lbl">ERROR</div><div class="stat-num">{c_err}</div></div>', unsafe_allow_html=True)

        # 테이블
        rows_data = []
        for p in plans:
            rows_data.append({
                "타입": p.kind,
                "행": p.row,
                "계정": p.account,
                "구분": p.action if not p.error else "오류",
                "계획 내용 및 검증 사유": p.preview()
            })
        df_plans = pd.DataFrame(rows_data)

        st.markdown("<div style='height: 16px;'></div>", unsafe_allow_html=True)
        st.dataframe(df_plans, use_container_width=True, hide_index=True)

        preview_xlsx_path = os.path.join(config.OUTPUT_DIR, f"{stem_name}_미리보기.xlsx")
        if os.path.exists(preview_xlsx_path):
            with open(preview_xlsx_path, "rb") as pf:
                st.download_button(
                    label="시뮬레이션 사본 내려받기 (.xlsx)",
                    data=pf.read(),
                    file_name=f"{stem_name}_미리보기.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        # 실제 반영 구역
        st.markdown("<div style='height: 24px;'></div>", unsafe_allow_html=True)
        st.markdown("""
        <div class="bat-card" style="border-left: 5px solid #000000;">
            <div class="bat-card-title">네이버 실서버 반영 (EXECUTE)</div>
            <div class="bat-card-desc">
                시뮬레이션 결과를 확인하였으며, 위 변경 내역을 네이버 광고시스템에 영구 반영합니다.
            </div>
        </div>
        """, unsafe_allow_html=True)

        confirm_check = st.checkbox("계획된 작업 내역을 최종 검토하였으며, 네이버 광고 계정에 실제 반영을 승인합니다.")

        if st.button("네이버 반영 실행", type="primary", disabled=not confirm_check):
            exec_logs = []
            def exec_log(msg):
                exec_logs.append(str(msg))

            progress_box = st.status("네이버 검색광고 API 반영 작업 진행 중...", expanded=True)
            with progress_box:
                try:
                    res = runner.apply(
                        st.session_state["last_file"],
                        env=ENV,
                        log=exec_log,
                        confirm=lambda summary: True
                    )
                    progress_box.update(label="반영 작업 완료", state="complete", expanded=False)
                    st.success(f"처리 결과: 성공 {res.get('done', 0)}건 / 실패 {res.get('failed', 0)}건 / 사전 오류 {res.get('errors', 0)}건")

                    result_files = sorted(
                        [os.path.join(config.OUTPUT_DIR, f) for f in os.listdir(config.OUTPUT_DIR)
                         if f.startswith(f"{stem_name}_결과_") and f.endswith(".xlsx")],
                        key=os.path.getmtime,
                        reverse=True
                    )
                    if result_files:
                        with open(result_files[0], "rb") as rf:
                            st.download_button(
                                label="최종 처리 결과 엑셀 내려받기 (.xlsx)",
                                data=rf.read(),
                                file_name=os.path.basename(result_files[0]),
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                type="primary"
                            )
                except Exception as e:
                    progress_box.update(label="작업 중 오류 중단", state="error", expanded=True)
                    st.error(f"오류: {e}")

            with st.expander("실행 상세 로그"):
                st.code("\n".join(exec_logs))


# ---------------------------------------------------------
# TAB 2: 현황 스냅샷
# ---------------------------------------------------------
with tab_snapshot:
    st.markdown("""
    <div class="bat-card">
        <div class="bat-card-title">광고 현황 데이터베이스 스냅샷</div>
        <div class="bat-card-desc">
            선택한 광고 계정의 최신 파워링크 키워드, 소재, 확장소재 6종의 전체 현황 및 고유 ID를 엑셀 파일로 추출합니다.
        </div>
    </div>
    """, unsafe_allow_html=True)

    if not AVAILABLE_ACCOUNTS:
        st.warning("조회 가능한 계정 키가 등록되지 않았습니다.")
    else:
        target_accounts = st.multiselect(
            "대상 광고 계정 선택",
            options=AVAILABLE_ACCOUNTS,
            default=AVAILABLE_ACCOUNTS[:1]
        )

        if st.button("현황 엑셀 생성", type="primary"):
            if not target_accounts:
                st.warning("계정을 1개 이상 선택하십시오.")
            else:
                snap_logs = []
                def snap_log(msg):
                    snap_logs.append(str(msg))

                snap_status = st.status("네이버 데이터 수집 및 현황 엑셀 생성 중...", expanded=True)
                with snap_status:
                    try:
                        out_path = runner.snapshot(target_accounts, env=ENV, log=snap_log)
                        if out_path and os.path.exists(out_path):
                            snap_status.update(label="현황 데이터 추출 완료", state="complete", expanded=False)
                            st.success("현황 파일 생성이 완료되었습니다.")
                            with open(out_path, "rb") as sf:
                                st.download_button(
                                    label="현황 엑셀 내려받기 (.xlsx)",
                                    data=sf.read(),
                                    file_name=os.path.basename(out_path),
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    type="primary"
                                )
                        else:
                            snap_status.update(label="현황 추출 실패", state="error", expanded=True)
                            st.warning("데이터가 없거나 파일 생성에 실패했습니다.")
                    except Exception as e:
                        snap_status.update(label="수집 중 오류 발생", state="error", expanded=True)
                        st.error(f"오류: {e}")

                with st.expander("수집 상세 로그"):
                    st.code("\n".join(snap_logs))
