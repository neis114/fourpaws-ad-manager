# BAT Performance AD OPS — 네이버 검색광고 일괄관리 웹 도구

포포즈(PETDOC) 및 담당 브랜드의 네이버 검색광고(파워링크)를 웹에서 안전하게 대량 제어(등록·수정·삭제)하고, 계정 전체 현황을 실시간 추출할 수 있는 **운영 자동화 웹 대시보드**입니다.

---

## 1. 프로젝트 개요 (Overview)
* **목적**: 네이버 대량 관리의 불편함을 개선하여, 단일 엑셀 양식에서 8대 대상(키워드·소재·확장소재 6종)의 등록·수정·삭제 작업을 웹 인터페이스로 안전하게 시뮬레이션 및 일괄 반영.
* **배포 형태**: Streamlit 기반 웹 애플리케이션 (사내 접속 비밀번호 인증)
* **담당 조직**: BAT 퍼포먼스 마케팅 포포즈 TF

---

## 2. 보안 아키텍처 및 안전성 (Security & Safety)

사내 보안 가이드라인을 준수하여 다음과 같은 보안 설계가 적용되어 있습니다.

### ① 자격 증명(API Key) 완전 격리 (No Hardcoding)
* 소스코드 내부 어디에도 네이버 API Key, Secret Key, Customer ID가 하드코딩되어 있지 않습니다.
* 모든 자격 증명은 **클라우드 환경변수(`Streamlit Secrets`)**에 암호화되어 안전하게 보관되며, 사용자 브라우저나 외부로 노출되지 않습니다.

### ② 무저장 구조 (Stateless & No-DB)
* 별도의 데이터베이스(RDBMS, NoSQL)를 구축하거나 연동하지 않습니다.
* 고객 개인정보(PII)를 일절 수집·저장하지 않으며, 순수 광고 키워드·문안 데이터만 엑셀 업로드 즉시 처리 후 반환하는 완전한 Stateless 구조입니다.

### ③ 접근 통제 (Authentication)
* 웹 애플리케이션 진입 시 **`APP_PASSWORD` 기반 세션 인증**이 필수로 요구되어, 외부인의 무단 접근 및 오작동을 차단합니다.

### ④ 파괴적 작업 방지 가드레일 (Safety Guardrails)
* **사전 시뮬레이션 (Dry-run)**: 네이버 광고 계정에 바로 반영되지 않으며, 문법 오류, 중복 여부, 필수값 누락을 먼저 검증하여 `[등록 N · 수정 N · 삭제 N · 오류 N]` 요약을 출력합니다.
* **반영 전 스냅샷 백업**: 실제 API 반영 직전 대상 항목들의 기존 상태를 백업(JSON)으로 자동 기록합니다.
* **최종 확인 다이얼로그**: 사용자가 모달 팝업 확인창에서 최종 승인 버튼을 명시적으로 클릭해야만 API 호출이 시작됩니다.

### ⑤ `.gitignore` 정보 유출 차단
* `.env`, `backup/`, `output/`, 로컬 작업 파일들이 GitHub 저장소에 커밋되지 않도록 배제되어 있습니다.

---

## 3. 핵심 기능 (Features)

### 01. 광고 일괄 관리 (Batch Operations)
* **지원 시트 8종**: 키워드, 단일형 소재, 확장소재 6종(추가제목, 추가설명, 홍보문구, 서브링크, 이미지형서브링크, 파워링크이미지)
* **표준 템플릿 제공**: 웹에서 즉시 최신 규격의 엑셀 템플릿(`.xlsx`) 다운로드
* **결과 리포트 자동 생성**: 처리 완료 후 행별 완료/실패 사유가 기록된 결과 사본 엑셀 제공

### 02. 계정 정보 내려받기 (Account Data)
* 광고 계정(PETDOC, BS 등)의 전체 키워드, 소재, 확장소재 및 고유 시스템 ID를 실시간 수집하여 엑셀로 추출
* 추출된 시스템 ID는 신규 작업 템플릿 작성 시 정확한 대상 지정용으로 활용

---

## 4. 디렉터리 구조 (Directory Structure)

```text
├── app.py                  # Streamlit 메인 웹 애플리케이션 (BAT B&W UI)
├── requirements.txt        # 클라우드 배포용 의존성 패키지 목록
├── .gitignore              # 민감 파일 및 로컬 산출물 커밋 차단 설정
├── 템플릿/                 # 표준 작업 엑셀 템플릿
├── ads_core/
│   ├── config.py           # 환경변수 로더 및 계정 파서
│   ├── naver_api.py        # 네이버 검색광고 공식 REST API 클라이언트 (HMAC 서명)
│   ├── specs.py            # 8대 시트 규격 및 유효성 검증 단일 기준
│   ├── planner.py          # 실행 계획 수립 및 문법/중복 사전 검증
│   ├── ext_planner.py      # 확장소재 6종 전용 계획 수립
│   ├── runner.py           # 시뮬레이션(preview) 및 실행(apply), 백업 파이프라인
│   └── excel_io.py         # 엑셀 I/O 및 결과 사본 기록
├── input/                  # 로컬 작업 엑셀 임시 저장소 (.gitignore 대상)
├── output/                 # 시뮬레이션 및 실행 결과 엑셀 (.gitignore 대상)
└── backup/                 # 반영 직전 스냅샷 백업 JSON (.gitignore 대상)
```

---

## 5. 기술 스택 (Tech Stack)

* **Language**: Python 3.10+
* **Framework**: Streamlit
* **Data Processing**: Pandas, OpenPyXL
* **Integration**: Naver Searchad REST API (HMAC-SHA256 Signature)
* **HTTP Client**: Requests

---

## 6. 환경변수 설정 명세 (Secrets)

Streamlit Cloud 배포 시 `Settings → Secrets`에 등록되는 키 명세입니다:

```toml
# 웹 접속 비밀번호
APP_PASSWORD = "..."

# 네이버 검색광고 API 자격 증명 (계정별)
NAVER_PETDOC_CUSTOMER_ID = "..."
NAVER_PETDOC_API_KEY = "..."
NAVER_PETDOC_SECRET_KEY = "..."

# 추가 계정 필요 시 동일 규격으로 등록 (NAVER_{ACCOUNT}_CUSTOMER_ID)
```
