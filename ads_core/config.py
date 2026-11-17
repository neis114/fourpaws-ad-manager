"""설정: .env 로딩, 작업 시트 열 구성."""

import os

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BACKUP_DIR = os.path.join(PROJECT_ROOT, 'backup')    # 삭제·수정 전 값(JSON)
INPUT_DIR = os.path.join(PROJECT_ROOT, 'input')      # 작업 엑셀을 올려 두는 곳
OUTPUT_DIR = os.path.join(PROJECT_ROOT, 'output')    # 결과가 적힌 사본 엑셀·현황 엑셀

ENV_CANDIDATES = [
    os.environ.get('ADS_MANAGER_ENV', ''),
    os.path.join(PROJECT_ROOT, '.env'),
    r'C:\포포즈\.env',
    os.path.join(os.path.expanduser('~'), '.naver-bs', '.env'),
]

# 작업 한 행의 내부 항목 이름. 실제 시트는 "키워드" / "소재" 시트로 나뉘어 있고(열 이름은 apps_script/Code.gs 의
# CONFIG.SHEETS), Apps Script 가 시트별 열("기존 키워드", "새 제목" 등)을 아래 이름으로 바꿔 보낸다. '대상'은 시트 이름이다.
#   기본은 이름(계정·캠페인·광고그룹·기존 값)으로 찾고, ID 열은 선택(적으면 ID 가 우선)이다.
COLS = ['계정', '캠페인', '광고그룹', '대상', '작업', '기존 값', '새 값',
        '입찰가', '소재 설명', 'PC URL', '모바일 URL', '상태',
        '광고그룹 ID(선택)', '대상 ID(선택)', '결과', '처리일시']
TARGETS = ('키워드', '소재')
ACTIONS = ('등록', '수정', '삭제')      # 교체는 없다: 값을 바꾸려면 등록 줄 + 삭제 줄

# 파워링크 텍스트 소재의 글자 수 안내 (넘으면 계획 단계에서 경고만 한다. 최종 판단은 네이버 API)
HEADLINE_MAX = 15
DESCRIPTION_MAX = 45


def list_inputs():
    """input 폴더의 작업 엑셀(.xlsx). 최근에 고친 것부터. 엑셀이 열어 둘 때 생기는 임시 잠금 파일(~$)은 뺀다."""
    if not os.path.isdir(INPUT_DIR):
        return []
    files = [os.path.join(INPUT_DIR, f) for f in os.listdir(INPUT_DIR)
             if f.lower().endswith('.xlsx') and not f.startswith('~$')]
    return sorted(files, key=os.path.getmtime, reverse=True)


def find_env_path():
    for p in ENV_CANDIDATES:
        if p and os.path.isfile(p):
            return p
    return None


def load_env(path=None):
    """.env 를 읽어 dict 로 반환한다 (os.environ 은 건드리지 않는다)."""
    path = path or find_env_path()
    if not path:
        raise FileNotFoundError('.env 파일을 찾지 못했습니다. 찾은 위치: ' + ', '.join(p for p in ENV_CANDIDATES if p))
    env = {}
    with open(path, encoding='utf-8-sig') as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith('#') or '=' not in line:
                continue
            k, v = line.split('=', 1)
            env[k.strip()] = v.strip().strip('"').strip("'")
    env['_PATH'] = path
    return env


# .env 의 계정 이름이 바뀌어도 동작하도록 새 이름 ↔ 옛 이름을 모두 찾아본다 (모니터링 자동화와 같은 규칙)
ACCOUNT_ALIASES = {
    'CAPITAL': ['PETDOC'], 'BUSAN': ['BS'], 'SEJONG': ['SJ'],
    'PETDOC': ['CAPITAL'], 'BS': ['BUSAN'], 'SJ': ['SEJONG'],
}


def account_keys(env, code):
    """계정 코드(PETDOC 등)의 (고객ID, API키, 비밀키). 없으면 None."""
    code = code.strip().upper()
    for c in [code] + ACCOUNT_ALIASES.get(code, []):
        vals = [env.get(f'NAVER_{c}_{k}', '') for k in ('CUSTOMER_ID', 'API_KEY', 'SECRET_KEY')]
        if all(vals):
            return tuple(vals)
    return None


def account_codes(env):
    """.env 에 키가 모두 있는 계정 코드 목록."""
    codes = []
    for k in env:
        if k.startswith('NAVER_') and k.endswith('_CUSTOMER_ID'):
            code = k[len('NAVER_'):-len('_CUSTOMER_ID')]
            if account_keys(env, code):
                codes.append(code)
    return codes
