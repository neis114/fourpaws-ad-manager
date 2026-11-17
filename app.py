"""네이버 검색광고 파워링크 일괄관리 웹 애플리케이션 (Streamlit)
- 클라우드(Streamlit Community Cloud) 및 사내 로컬 브라우저 지원
- 접속 비밀번호 인증 및 API 키 격리 관리
"""

import io
import os
import sys
import tempfile
import pandas as pd
import streamlit as st

# 현재 디렉터리를 sys.path에 추가하여 ads_core 모듈 import
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
if CURRENT_DIR not in sys.path:
    sys.path.insert(0, CURRENT_DIR)

from ads_core import config, runner, specs  # noqa: E402
from ads_core.excel_io import read_jobs  # noqa: E402
from ads_core.planner import Planner  # noqa: E402


# ---------------------------------------------------------
# 1. 페이지 설정 & 테마
# ---------------------------------------------------------
st.set_page_config(
    page_title="네이버 광고 일괄관리",
    page_icon="📢",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    .metric-card {
        background-color: #f8f9fa;
        border-radius: 8px;
        padding: 15px;
        border-left: 5px solid #03c75a;
        margin-bottom: 10px;
    }
    .stButton>button {
        border-radius: 6px;
        font-weight: 600;
    }
</style>
""", unsafe_allow_html=True)


# ---------------------------------------------------------
# 2. 보안 인증 (Password Protection)
# ---------------------------------------------------------
def check_auth():
    """Secrets 에 APP_PASSWORD 가 지정되어 있으면 비밀번호 입력을 요구합니다."""
    expected_pw = None
    if hasattr(st, "secrets") and "APP_PASSWORD" in st.secrets:
        expected_pw = str(st.secrets["APP_PASSWORD"]).strip()

    if not expected_pw:
        return True  # 비밀번호 설정이 없으면 인증 통과 (로컬 개발용)

    if st.session_state.get("authenticated", False):
        return True

    st.title("🔒 로그인")
    st.write("사내 보안을 위해 접근 비밀번호를 입력해 주세요.")
    
    with st.form("login_form"):
        input_pw = st.text_input("접속 비밀번호", type="password")
        submit = st.form_submit_button("접속하기")
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
    """Streamlit Secrets 또는 로컬 .env 에서 네이버 API 키를 dict 형태로 취합합니다."""
    env = {}

    # 1) Streamlit secrets 우선 로드 (클라우드 배포 환경)
    if hasattr(st, "secrets"):
        for k, v in st.secrets.items():
            if k.startswith("NAVER_"):
                env[k] = str(v).strip().strip('"').strip("'")

    # 2) Secrets에 키가 없으면 로컬 .env 파일 검색 (로컬 실행 환경)
    if not env:
        try:
            env = config.load_env()
        except Exception:
            env = {}

    return env


ENV = get_env_dict()
AVAILABLE_ACCOUNTS = config.account_codes(ENV)


# ---------------------------------------------------------
# 4. 사이드바: 계정 현황 및 바로가기
# ---------------------------------------------------------
with st.sidebar:
    st.title("📢 네이버 광고 관리")
    st.caption("파워링크 키워드 · 소재 · 확장소재 일괄 등록/수정/삭제")
    
    st.divider()
    st.subheader("🔑 연결된 광고 계정")
    if AVAILABLE_ACCOUNTS:
        for acc in AVAILABLE_ACCOUNTS:
            st.success(f"계정: {acc}")
    else:
        st.warning("등록된 계정 API 키가 없습니다.\nStreamlit Secrets 나 .env 설정을 확인하세요.")

    st.divider()
    st.subheader("📄 템플릿 다운로드")
    template_path = os.path.join(CURRENT_DIR, "템플릿", "광고_일괄관리_템플릿_확장소재포함.xlsx")
    if os.path.exists(template_path):
        with open(template_path, "rb") as f:
            st.download_button(
                label="📥 최신 엑셀 템플릿 받기",
                data=f.read(),
                file_name="광고_일괄관리_템플릿_확장소재포함.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
    
    if st.session_state.get("authenticated", False) and hasattr(st, "secrets") and "APP_PASSWORD" in st.secrets:
        st.divider()
        if st.button("🚪 로그아웃", use_container_width=True):
            st.session_state["authenticated"] = False
            st.rerun()


# ---------------------------------------------------------
# 5. 메인 화면 탭 구성
# ---------------------------------------------------------
tab_manage, tab_snapshot = st.tabs(["🚀 광고 일괄관리 (업로드/실행)", "📊 광고 현황 가져오기 (Snapshot)"])


# ---------------------------------------------------------
# TAB 1: 광고 일괄관리
# ---------------------------------------------------------
with tab_manage:
    st.subheader("1. 작업 엑셀 업로드")
    st.write("템플릿에 키워드/소재/확장소재 작업 내역을 작성한 엑셀 파일(.xlsx)을 업로드하세요.")
    
    uploaded_file = st.file_uploader(
        "작업 엑셀 파일 선택",
        type=["xlsx"],
        help="대상별 시트(키워드, 소재, 추가제목 등)에 작업(등록/수정/삭제)을 작성한 파일"
    )

    if uploaded_file is not None:
        # 업로드된 파일을 임시 디렉터리에 저장하여 기존 모듈과 연동
        temp_dir = tempfile.mkdtemp()
        temp_input_path = os.path.join(temp_dir, uploaded_file.name)
        with open(temp_input_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        st.success(f"파일 업로드 완료: `{uploaded_file.name}`")

        # -------------------------------------------------
        # 단계 A: 미리보기(계획 수립) 실행
        # -------------------------------------------------
        st.divider()
        st.subheader("2. 작업 계획 미리보기 (Dry-Run)")
        st.caption("네이버 광고에는 아무런 변경을 가하지 않고, 엑셀 문법과 대상 ID를 조회하여 실행 계획을 검증합니다.")

        col_btn1, _ = st.columns([2, 5])
        with col_btn1:
            btn_preview = st.button("🔍 변경 계획 검증하기 (미리보기)", use_container_width=True, type="primary")

        if btn_preview or st.session_state.get("last_preview_file") == uploaded_file.name:
            logs = []
            def log_fn(msg):
                logs.append(msg)

            with st.spinner("계획을 수립하고 네이버 계정 대상을 조회 중입니다..."):
                try:
                    plans = runner.preview(temp_input_path, env=ENV, log=log_fn)
                    st.session_state["last_preview_file"] = uploaded_file.name
                    st.session_state["last_preview_plans"] = plans
                    st.session_state["temp_input_path"] = temp_input_path
                except Exception as e:
                    st.error(f"미리보기 중 오류가 발생했습니다: {e}")
                    with st.expander("세부 로그 보기"):
                        st.code("\n".join(logs))
                    st.stop()

            # 카운트 요약
            plans = st.session_state["last_preview_plans"]
            counts = runner._counts(plans)

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("등록 예정", f"{counts.get('등록', 0)} 건")
            c2.metric("수정 예정", f"{counts.get('수정', 0)} 건")
            c3.metric("삭제 예정", f"{counts.get('삭제', 0)} 건")
            err_cnt = counts.get('오류', 0)
            c4.metric("사전 오류", f"{err_cnt} 건", delta=f"-{err_cnt}" if err_cnt > 0 else None, delta_color="inverse")

            # 상세 내역 테이블
            plan_records = []
            for p in plans:
                plan_records.append({
                    "시트": p.kind,
                    "행": p.row,
                    "작업": p.action,
                    "계정": p.account,
                    "대상": getattr(p, "target_text", "") or getattr(p, "name", ""),
                    "계획 요약": p.preview(),
                    "오류 여부": "⚠️ 오류" if p.error else "정상"
                })
            df_plans = pd.DataFrame(plan_records)
            
            st.dataframe(
                df_plans,
                column_config={
                    "오류 여부": st.column_config.TextColumn(
                        "상태",
                        help="오류가 있으면 실제 실행 시 해당 행은 건너뛰어집니다."
                    )
                },
                use_container_width=True,
                height=300
            )

            # 미리보기 사본 엑셀 다운로드
            preview_output_path = os.path.join(config.OUTPUT_DIR, f"{runner._stem(temp_input_path)}_미리보기.xlsx")
            if os.path.exists(preview_output_path):
                with open(preview_output_path, "rb") as pf:
                    st.download_button(
                        label="📥 [미리보기 결과 사본 엑셀] 다운로드",
                        data=pf.read(),
                        file_name=os.path.basename(preview_output_path),
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

            # -------------------------------------------------
            # 단계 B: 네이버 실제 반영 (Apply)
            # -------------------------------------------------
            st.divider()
            st.subheader("3. 네이버 광고 실제 반영")
            st.warning("⚠️ 실행 버튼을 누르면 실제 네이버 검색광고 계정에 키워드/소재/입찰가 변경사항이 적용됩니다.")

            confirm_check = st.checkbox("위 계획 요약 및 내역을 확인했으며, 실제 네이버 광고에 반영하는 것에 동의합니다.")

            if st.button("🚀 네이버 광고 계정에 반영 실행", disabled=not confirm_check, type="secondary"):
                run_logs = []
                log_box = st.empty()

                def run_log_fn(msg):
                    run_logs.append(msg)
                    log_box.code("\n".join(run_logs[-20:]), language="text")

                with st.spinner("네이버 API를 호출하여 광고를 일괄 변경 중입니다..."):
                    try:
                        res = runner.apply(
                            st.session_state["temp_input_path"],
                            env=ENV,
                            log=run_log_fn,
                            confirm=lambda _: True  # 웹 UI 체크박스로 이미 확인 완료
                        )
                        st.success(f"반영 완료! (성공: {res.get('done', 0)}건, 실패: {res.get('failed', 0)}건, 오류 제외: {res.get('errors', 0)}건)")

                        # 결과 사본 엑셀 찾기
                        import glob
                        stem = runner._stem(st.session_state["temp_input_path"])
                        results_files = sorted(glob.glob(os.path.join(config.OUTPUT_DIR, f"{stem}_결과_*.xlsx")), reverse=True)
                        if results_files:
                            with open(results_files[0], "rb") as rf:
                                st.download_button(
                                    label="🎉 최종 실행 결과 엑셀 다운로드",
                                    data=rf.read(),
                                    file_name=os.path.basename(results_files[0]),
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    type="primary"
                                )
                    except Exception as e:
                        st.error(f"실행 중 치명적 오류 발생: {e}")


# ---------------------------------------------------------
# TAB 2: 현황 가져오기 (Snapshot)
# ---------------------------------------------------------
with tab_snapshot:
    st.subheader("계정별 광고 현황 다운로드 (스냅샷)")
    st.write("선택한 계정의 모든 파워링크 키워드, 소재, 확장소재의 현재 상태와 고유 ID를 엑셀 파일로 추출합니다.")

    if not AVAILABLE_ACCOUNTS:
        st.error("사용 가능한 네이버 계정 정보가 없습니다.")
    else:
        selected_accounts = st.multiselect(
            "현황을 가져올 계정 선택",
            options=AVAILABLE_ACCOUNTS,
            default=AVAILABLE_ACCOUNTS[:1]
        )

        if st.button("📊 현황 엑셀 생성 시작", disabled=not selected_accounts, type="primary"):
            snap_logs = []
            snap_log_box = st.empty()

            def snap_log(msg):
                snap_logs.append(msg)
                snap_log_box.code("\n".join(snap_logs[-15:]), language="text")

            with st.spinner("네이버에서 전체 캠페인/광고그룹/키워드/소재 데이터를 수집 중입니다..."):
                try:
                    out_path = runner.snapshot(selected_accounts, env=ENV, log=snap_log)
                    if out_path and os.path.exists(out_path):
                        st.success(f"현황 수집 완료! 파일: `{os.path.basename(out_path)}`")
                        with open(out_path, "rb") as sf:
                            st.download_button(
                                label="📥 [현황 엑셀] 지금 다운로드",
                                data=sf.read(),
                                file_name=os.path.basename(out_path),
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                type="primary"
                            )
                    else:
                        st.warning("수집된 데이터가 없습니다. (파워링크 캠페인이 없거나 조회 실패)")
                except Exception as e:
                    st.error(f"현황 가져오기 실패: {e}")
