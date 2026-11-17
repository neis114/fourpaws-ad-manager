"""시트 구성의 단일 기준: 시트별 열 이름, 확장소재 종류별 값 열·API 변환.

템플릿 엑셀 생성, 시트 읽기(runner), 확장소재 계획(ext_planner)이 모두 이 파일을 쓴다.
구글시트 쪽 Apps Script 는 열 구성을 모르고 머리글 이름대로 읽어 오기만 한다.
"""

import datetime as dt
import re

TAIL = ['결과', '처리일시']
HEAD = ['작업', '계정', '캠페인', '광고그룹']   # 작업은 맨 앞에 두어 계정·캠페인·광고그룹·기존 값 열이 붙어 있게 한다(순서는 코드와 무관, 이름으로 찾는다)
PERIOD_COLS = ['기간 시작일', '기간 종료일']
KST = dt.timezone(dt.timedelta(hours=9))


class SpecError(ValueError):
    """입력 값을 API 형식으로 바꿀 수 없다(사유가 메시지)."""


# ── 확장소재 종류 ────────────────────────────────────────────────────
class ExtKind:
    """확장소재 한 종류. 시트 이름이 곧 대상 이름이다.

    key_col 이 이 종류의 "기본 값 열"(추가제목, 링크1 이름, 이미지 ID …)이다. 등록에서는 새로 만들 값이고,
    수정·삭제에서는 기존 항목을 찾는 값이다(확장소재 ID 를 적으면 ID 가 우선). 값 열을 바꾸는 "교체"는 없다.

    build(vals, ctx) -> adExtension(API 형식)   vals: {값 열 이름: 문자열}
    read(adExtension) -> {값 열 이름: 문자열}   현재 값을 시트 표기로 (중복·현황에 쓴다)
    match(text, adExtension) -> bool          기본 값 열의 값이 이 확장소재를 가리키는가
    label(adExtension) -> str                 로그에 보일 짧은 이름
    """

    def __init__(self, name, type_, values, build, read, match, label, key_col=None):
        self.name, self.type, self.values = name, type_, values
        self.build, self.read, self.match, self.label = build, read, match, label
        self.key_col = key_col or values[0]      # 찾는 값이 들어가는 열(기본은 첫 값 열)

    @property
    def cols(self):
        return HEAD + self.values + PERIOD_COLS + ['상태', '광고그룹 ID(선택)', '확장소재 ID(선택)'] + TAIL

    @property
    def snapshot_header(self):
        return ['계정', '캠페인', '광고그룹'] + self.values + PERIOD_COLS + ['상태', '검수상태', '광고그룹 ID', '확장소재 ID']


def _s(v):
    return (v or '').strip()


# 추가제목 (HEADLINE): 노출 위치 pin = 1 | 2 | 없음(모든 위치)
POS_LABELS = {None: '모든 위치에 노출 가능', 1: '1번째 위치에만 노출', 2: '2번째 위치에만 노출'}
POS_OPTIONS = list(POS_LABELS.values())


def _pin(text):
    t = _s(text)
    if not t or '모든' in t:
        return None
    m = re.search(r'[12]', t)
    if not m:
        raise SpecError(f'노출 가능 위치는 {" / ".join(POS_OPTIONS)} 중에서 골라주세요: "{text}"')
    return int(m.group())


def _build_headline(v, ctx):
    head = _s(v['추가제목'])
    if not head:
        raise SpecError('추가제목을 적어주세요')
    ad = {'headline': head}
    pin = _pin(v['노출 가능 위치 지정'])
    if pin:
        ad['pin'] = pin
    return ad


# 홍보문구 (DESCRIPTION): 유형 heading 은 비우거나 "선택 안 함"이면 보내지 않는다
NO_HEADING = '선택 안 함'


def _build_promo(v, ctx):
    desc = _s(v['홍보문구'])
    if not desc:
        raise SpecError('홍보문구를 적어주세요')
    ad = {'description': desc}
    head = _s(v['유형'])
    if head and head != NO_HEADING:
        ad['heading'] = head
    return ad


def _build_desc_extra(v, ctx):
    desc = _s(v['추가설명'])
    if not desc:
        raise SpecError('추가설명을 적어주세요')
    return {'description': desc}


# 서브링크 (SUB_LINKS): 링크 3~4개 {name, final}
def _build_sublinks(v, ctx):
    links = []
    for i in range(1, 5):
        name, url = _s(v[f'링크{i} 이름']), _s(v[f'링크{i} URL'])
        if bool(name) != bool(url):
            raise SpecError(f'링크{i}의 이름과 URL을 모두 적어주세요')
        if name:
            links.append({'name': name, 'final': url})
    if len(links) < 3:
        raise SpecError('서브링크는 링크를 3개 이상 적어주세요')
    return links


# 이미지형 서브링크 (IMAGE_SUB_LINKS): 링크 3개 {imagePath, name, final}
def _build_image_sublinks(v, ctx):
    links = []
    for i in range(1, 4):
        img, name, url = _s(v[f'링크{i} 이미지 ID']), _s(v[f'링크{i} 이름']), _s(v[f'링크{i} URL'])
        if not (img and name and url):
            raise SpecError(f'링크{i}의 이미지 ID·이름·URL을 모두 적어주세요 (이미지형 서브링크는 3개 필요)')
        links.append({'imagePath': ctx.image(img)['imagePath'], 'name': name, 'final': url})
    return links


# 파워링크 이미지 (POWER_LINK_IMAGE): {imagePath, imageWidth, imageHeight}
def _build_image(v, ctx):
    img = _s(v['이미지 ID'])
    if not img:
        raise SpecError('이미지 ID를 적어주세요')
    found = ctx.image(img)
    ad = {'imagePath': found['imagePath']}
    if found.get('imageWidth'):
        ad.update({'imageWidth': found['imageWidth'], 'imageHeight': found['imageHeight']})
    return ad


def _read_headline(a):
    return {'추가제목': a.get('headline', ''), '노출 가능 위치 지정': POS_LABELS.get(a.get('pin'), str(a.get('pin')))}


def _read_promo(a):
    return {'홍보문구': a.get('description', ''), '유형': a.get('heading') or NO_HEADING}


def _read_sublinks(a):
    out = {}
    for i in range(1, 5):
        link = a[i - 1] if i <= len(a) else {}
        out[f'링크{i} 이름'], out[f'링크{i} URL'] = link.get('name', ''), link.get('final', '')
    return out


def _read_image_sublinks(a):
    out = {}
    for i in range(1, 4):
        link = a[i - 1] if i <= len(a) else {}
        out[f'링크{i} 이미지 ID'], out[f'링크{i} 이름'], out[f'링크{i} URL'] = (
            link.get('imagePath', ''), link.get('name', ''), link.get('final', ''))
    return out


def _text_match(key):
    return lambda text, a: _s(text) == _s((a or {}).get(key))


def _first_link_name(text, a):
    return bool(a) and _s(text) == _s(a[0].get('name'))


def _image_match(text, a):
    t, path = _s(text), (a or {}).get('imagePath', '')
    return bool(t) and (t == path or t in path)


def _sublink_cols(n, with_image=False):
    cols = []
    for i in range(1, n + 1):
        cols += ([f'링크{i} 이미지 ID'] if with_image else []) + [f'링크{i} 이름', f'링크{i} URL']
    return cols


EXT_KINDS = [
    ExtKind('추가제목', 'HEADLINE', ['추가제목', '노출 가능 위치 지정'],
            _build_headline, _read_headline, _text_match('headline'), lambda a: a.get('headline', '')),
    ExtKind('추가설명', 'DESCRIPTION_EXTRA', ['추가설명'],
            _build_desc_extra, lambda a: {'추가설명': a.get('description', '')}, _text_match('description'),
            lambda a: a.get('description', '')),
    ExtKind('홍보문구', 'DESCRIPTION', ['홍보문구', '유형'],
            _build_promo, _read_promo, _text_match('description'), lambda a: a.get('description', '')),
    ExtKind('서브링크', 'SUB_LINKS', _sublink_cols(4),
            _build_sublinks, _read_sublinks, _first_link_name, lambda a: ' / '.join(x['name'] for x in a)),
    ExtKind('이미지형서브링크', 'IMAGE_SUB_LINKS', _sublink_cols(3, with_image=True),
            _build_image_sublinks, _read_image_sublinks, _first_link_name, lambda a: ' / '.join(x['name'] for x in a),
            key_col='링크1 이름'),
    ExtKind('파워링크이미지', 'POWER_LINK_IMAGE', ['이미지 ID'],
            _build_image, lambda a: {'이미지 ID': a.get('imagePath', '')}, _image_match,
            lambda a: '…' + (a.get('imagePath', '')[-24:])),
]
EXT = {k.name: k for k in EXT_KINDS}


# ── 시트 목록 ────────────────────────────────────────────────────────
# 키워드·소재는 Python 내부 항목 이름(planner 가 씀)으로 바꾸는 map 이 있다: {내부 이름: 시트 열 이름}
#   "키워드"/"제목" 칸은 등록에서는 새로 만들 값('새 값'), 수정·삭제에서는 찾는 값('기존 값')이라 두 내부 이름을 같은 열에 연결한다.
KEYWORD_COLS = HEAD + ['키워드', '입찰가', '상태', '광고그룹 ID(선택)', '키워드 ID(선택)'] + TAIL
AD_COLS = HEAD + ['제목', '소재 설명', 'PC URL', '모바일 URL', '상태', '광고그룹 ID(선택)', '소재 ID(선택)'] + TAIL

SHEETS = [
    {'name': '키워드', 'cols': KEYWORD_COLS,
     'map': {'기존 값': '키워드', '새 값': '키워드', '대상 ID(선택)': '키워드 ID(선택)'}},
    {'name': '소재', 'cols': AD_COLS,
     'map': {'기존 값': '제목', '새 값': '제목', '대상 ID(선택)': '소재 ID(선택)'}},
] + [{'name': k.name, 'cols': k.cols, 'map': {'대상 ID(선택)': '확장소재 ID(선택)'}} for k in EXT_KINDS]
BY_SHEET = {s['name']: s for s in SHEETS}
KINDS = tuple(BY_SHEET)           # 대상 이름 = 시트 이름
ID_COLS = ('광고그룹 ID(선택)', '키워드 ID(선택)', '소재 ID(선택)', '확장소재 ID(선택)')


# ── 열 구분: 필수 / 조건부 필수 / 선택 / 자동 ───────────────────────────
# 템플릿의 머리글 색·메모·"열 안내" 시트가 모두 여기서 나온다. 실제 검사는 planner·ext_planner 가 하며, 아래 설명은 그 동작과 같아야 한다.
REQUIRED, CONDITIONAL, OPTIONAL, AUTO = '필수', '조건부 필수', '선택', '자동'
LEVELS = (REQUIRED, CONDITIONAL, OPTIONAL, AUTO)

_GROUP_NOTE = '광고그룹 ID를 적지 않으면 필수. 광고주센터의 이름과 정확히 같아야 한다'
COMMON_INFO = {
    '계정': (REQUIRED, '광고계정 코드. 예: PETDOC, BS, SJ'),
    '캠페인': (CONDITIONAL, _GROUP_NOTE),
    '광고그룹': (CONDITIONAL, _GROUP_NOTE),
    '작업': (REQUIRED, '등록 / 수정 / 삭제 중 하나를 고른다. 값 자체를 바꾸려면 등록 줄과 삭제 줄 두 줄로 적는다(같은 값을 다시 등록하면 삭제가 먼저 처리된다)'),
    '상태': (OPTIONAL, 'ON / OFF. 등록에서 비우면 ON, 수정에서 비우면 기존 상태를 유지한다'),
    '기간 시작일': (OPTIONAL, '2026-10-01 형식. 시작일·종료일을 함께 적는다. 비우면 기간 없음(수정에서는 기존 기간 유지). 기간을 없애려면 시작일에 "해제"'),
    '기간 종료일': (OPTIONAL, '2026-10-31 형식. 시작일과 함께 적는다'),
    '광고그룹 ID(선택)': (OPTIONAL, '적으면 캠페인·광고그룹 이름 대신 이 ID로 광고그룹을 찾는다'),
    '결과': (AUTO, '입력하지 않는다. 미리보기·실행 결과가 output 폴더의 사본 파일에 자동으로 적힌다. "완료"인 행은 다시 실행하지 않는다'),
    '처리일시': (AUTO, '입력하지 않는다. 실행한 시각이 자동으로 적힌다'),
}
SHEET_INFO = {
    '키워드': {
        '키워드': (CONDITIONAL, '등록: 새로 만들 키워드(필수). 수정·삭제: 찾을 키워드(키워드 ID를 적으면 생략 가능). 키워드 문구는 수정할 수 없다'),
        '입찰가': (OPTIONAL, '등록: 비우면 광고그룹 입찰가. 수정: 입찰가·상태 중 하나는 적어야 한다'),
        '키워드 ID(선택)': (OPTIONAL, '적으면 키워드 칸 대신 이 ID로 수정·삭제할 대상을 찾는다'),
    },
    '소재': {
        '제목': (CONDITIONAL, '등록: 새로 만들 제목(필수). 수정·삭제: 찾을 제목(소재 ID를 적으면 생략 가능). 같은 제목이 여럿이면 소재 ID로 지정. 제목은 수정할 수 없다'),
        '소재 설명': (CONDITIONAL, '등록에서 필수. 수정에서는 바꿀 때만 적는다(비우면 기존 설명 유지)'),
        'PC URL': (CONDITIONAL, '등록에서 필수. 수정에서는 바꿀 때만 적는다'),
        '모바일 URL': (OPTIONAL, '등록에서 비우면 PC URL과 같게 등록된다. 수정에서는 바꿀 때만 적는다'),
        '소재 ID(선택)': (OPTIONAL, '적으면 제목 칸 대신 이 ID로 수정·삭제할 대상을 찾는다'),
    },
}
EXT_KEY_NOTE = '등록: 새로 만들 값(필수). 수정·삭제: 찾을 값(확장소재 ID를 적으면 생략 가능). 이 값 자체는 수정할 수 없다'
EXT_VALUE_DEFAULT = (CONDITIONAL, '등록에서 필수. 수정에서는 바꿀 값만 적는다(비우면 기존 값 유지)')
EXT_VALUE_INFO = {
    '노출 가능 위치 지정': (OPTIONAL, '비우면 모든 위치. 수정에서 비우면 기존 값 유지'),
    '유형': (OPTIONAL, '비우거나 "선택 안 함"이면 유형 없음(예: 이벤트). 수정에서 비우면 기존 값 유지'),
    '링크4 이름': (OPTIONAL, '서브링크는 3개 이상이면 되므로 4번째 링크는 선택'),
    '링크4 URL': (OPTIONAL, '링크4 이름과 함께 적는다'),
}


def col_info(sheet, col):
    """시트의 열 → (구분, 설명). 구분은 LEVELS 중 하나."""
    if col in SHEET_INFO.get(sheet, {}):
        return SHEET_INFO[sheet][col]
    kind = EXT.get(sheet)
    if kind:
        if col == kind.key_col:
            return CONDITIONAL, EXT_KEY_NOTE
        if col == '확장소재 ID(선택)':
            return OPTIONAL, f'적으면 "{kind.key_col}" 칸 대신 이 ID로 수정·삭제할 대상을 찾는다'
        if col in kind.values:
            return EXT_VALUE_INFO.get(col, EXT_VALUE_DEFAULT)
    return COMMON_INFO[col]


# ── 기간(KST 날짜 ↔ 네이버 UTC 일시) ─────────────────────────────────
_DATE_RE = re.compile(r'^\s*(\d{4})[-./\s]+(\d{1,2})[-./\s]+(\d{1,2})\.?\s*$')


def date_to_iso(text):
    """"2026-10-01" 등 → "2026-09-30T15:00:00.000Z" (한국 시간 0시를 네이버가 저장하는 UTC 형식으로)."""
    m = _DATE_RE.match(text or '')
    if not m:
        raise SpecError(f'날짜는 2026-10-01 형식으로 적어주세요: "{text}"')
    try:
        d = dt.datetime(int(m[1]), int(m[2]), int(m[3]), tzinfo=KST)
    except ValueError:
        raise SpecError(f'없는 날짜입니다: "{text}"')
    return d.astimezone(dt.timezone.utc).strftime('%Y-%m-%dT%H:%M:%S.000Z')


def iso_to_date(iso):
    if not iso:
        return ''
    t = dt.datetime.strptime(iso[:19], '%Y-%m-%dT%H:%M:%S').replace(tzinfo=dt.timezone.utc)
    return t.astimezone(KST).strftime('%Y-%m-%d')
