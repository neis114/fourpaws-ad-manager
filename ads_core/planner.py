"""작업 시트 한 행 → 실행 계획. 네이버의 현재 상태를 읽어 무엇을 할지(또는 왜 못 하는지)를 정한다.

대상은 이름(계정·캠페인·광고그룹·기존 값)으로 찾는다. ID 열이 채워져 있으면 ID 를 우선한다.
"""

from dataclasses import dataclass, field

from . import config, specs
from .naver_api import AD_TYPE, ApiError, NaverAdsApi, ad_type_label


class PlanError(Exception):
    """이 행은 실행할 수 없다(사유가 메시지)."""


@dataclass
class Op:
    desc: str
    run: object                 # 인자 없는 함수. 네이버 API 를 호출한다(성공하면 desc 가 결과 메시지가 된다)
    backup: dict = None         # 바뀌기 전의 값(삭제·수정 때만)


@dataclass
class Plan:
    row: int
    action: str = ''
    kind: str = ''
    account: str = ''
    ops: list = field(default_factory=list)
    error: str = ''
    warns: list = field(default_factory=list)

    def preview(self):
        if self.error:
            return f'오류: {self.error}'
        text = '예정: ' + ' → '.join(op.desc for op in self.ops)
        return text + (' / 주의: ' + '; '.join(self.warns) if self.warns else '')


def _kw_key(s):
    """키워드 비교용: 공백을 없애고 소문자로 맞춘다(네이버는 띄어쓰기·대소문자만 다른 키워드를 같은 것으로 본다)."""
    return ''.join(str(s).split()).lower()


def _status(value):
    """상태 칸 → userLock(True=OFF) 또는 None(변경 없음)."""
    v = value.strip().upper()
    if not v:
        return None
    if v in ('ON', '켜기', '노출'):
        return False
    if v in ('OFF', '끄기', '중지'):
        return True
    raise PlanError(f'상태는 ON 또는 OFF 로 적어주세요: "{value}"')


def _bid(value):
    v = value.replace(',', '').strip()
    if not v:
        return None
    try:
        n = int(float(v))
    except ValueError:
        raise PlanError(f'입찰가가 숫자가 아닙니다: "{value}"')
    if n <= 0:
        raise PlanError(f'입찰가는 0보다 커야 합니다: "{value}"')
    return n


class Planner:
    def __init__(self, env):
        self.env = env
        self._apis = {}
        self._groups = {}
        self._images = {}      # 고객 ID -> 이미지 목록(image_index)
        self._items = {}       # (고객 ID, 광고그룹 ID, 종류) -> 현재 목록. 같은 실행 안의 앞 행 결과를 뒤 행이 알 수 있게 계획하며 갱신한다

    # ── 조회 도우미 ──────────────────────────────────────────────────
    def api(self, account):
        code = account.strip().upper()
        if not code:
            raise PlanError('계정이 비어 있습니다')
        if code not in self._apis:
            keys = config.account_keys(self.env, code)
            if not keys:
                raise PlanError(f'.env 에 {code} 계정 키가 없습니다')
            self._apis[code] = NaverAdsApi(*keys)
        return self._apis[code]

    def _group(self, api, job, existing):
        gid = job['광고그룹 ID(선택)'] or (existing or {}).get('nccAdgroupId', '')
        if existing and job['광고그룹 ID(선택)'] and existing['nccAdgroupId'] != job['광고그룹 ID(선택)']:
            raise PlanError('대상 ID 가 지정한 광고그룹에 속하지 않습니다')
        key = (api.customer_id, gid or (job['캠페인'], job['광고그룹']))
        if key in self._groups:
            return self._groups[key]
        try:
            if gid:
                grp = api.get_adgroup(gid)
                if grp.get('adgroupType') != 'WEB_SITE':
                    raise PlanError('파워링크 광고그룹이 아닙니다')
            else:
                if not (job['캠페인'] and job['광고그룹']):
                    raise PlanError('캠페인·광고그룹(또는 광고그룹 ID)을 적어주세요')
                grp = api.find_adgroup(job['캠페인'], job['광고그룹'])
        except LookupError as e:
            raise PlanError(str(e))
        except ApiError as e:
            raise PlanError(f'광고그룹 조회 실패: {e}')
        self._groups[key] = grp
        return grp

    def _list(self, api, grp, kind):
        group = '확장소재' if kind in specs.EXT else kind      # 확장소재는 종류가 달라도 한 번에 조회해 함께 쓴다
        key = (api.customer_id, grp['nccAdgroupId'], group)
        if key not in self._items:
            fetch = {'키워드': api.keywords, '소재': api.ads, '확장소재': api.extensions}[group]
            self._items[key] = list(fetch(grp['nccAdgroupId']))
        return self._items[key]

    def image_index(self, api):
        """계정의 이미지 → {imagePath: {imagePath, imageWidth, imageHeight}}. 기존 이미지형 확장소재에서 모은다(계정당 1회)."""
        if api.customer_id not in self._images:
            index = {}
            for camp in api.powerlink_campaigns():
                for grp in api.adgroups(camp['nccCampaignId']):
                    for x in api.extensions(grp['nccAdgroupId']):
                        a = x.get('adExtension')
                        if x.get('type') == 'POWER_LINK_IMAGE' and a:
                            index[a['imagePath']] = {'imagePath': a['imagePath'], 'imageWidth': a.get('imageWidth'),
                                                     'imageHeight': a.get('imageHeight')}
                        elif x.get('type') == 'IMAGE_SUB_LINKS' and a:
                            for link in a:
                                index.setdefault(link['imagePath'], {'imagePath': link['imagePath']})
            self._images[api.customer_id] = index
        return self._images[api.customer_id]

    # ── 행 → 계획 ────────────────────────────────────────────────────
    def plan(self, job):
        plan = Plan(row=job['row'], action=job['작업'], kind=job['대상'], account=job['계정'].strip().upper())
        try:
            if job['대상'] not in specs.KINDS:
                raise PlanError(f'대상은 {" / ".join(specs.KINDS)} 중 하나여야 합니다')
            if job['작업'] == '교체':
                raise PlanError('교체는 지원하지 않습니다. 등록 줄과 삭제 줄 두 줄로 나눠 적어 주세요(다른 값으로 바꾸면 등록 줄을 위에 두면 노출 공백이 없고, 같은 값을 다시 등록하면 삭제가 자동으로 먼저 처리됩니다)')
            if job['작업'] not in config.ACTIONS:
                raise PlanError(f'작업은 {" / ".join(config.ACTIONS)} 중 하나여야 합니다')
            api = self.api(job['계정'])
            existing = self._by_id(api, job) if job['대상 ID(선택)'] else None
            grp = self._group(api, job, existing)
            items = self._list(api, grp, job['대상'])
            if job['대상'] in specs.EXT:
                from .ext_planner import plan_extension   # 순환 import 를 피하려고 여기서 불러온다
                plan_extension(self, plan, api, grp, items, job, existing, job['대상'])
            else:
                (self._keyword if job['대상'] == '키워드' else self._ad)(plan, api, grp, items, job, existing)
            if not plan.ops:
                raise PlanError('할 일이 없습니다')
        except PlanError as e:
            plan.error = str(e)
            plan.ops = []
        except ApiError as e:
            plan.error = f'조회 실패: {e}'
            plan.ops = []
        return plan

    def _by_id(self, api, job):
        getter = {'키워드': api.get_keyword, '소재': api.get_ad}.get(job['대상'], api.get_extension)
        try:
            return getter(job['대상 ID(선택)'])
        except ApiError:
            raise PlanError(f'대상 ID 를 찾지 못했습니다: {job["대상 ID(선택)"]}')

    # ── 키워드 ───────────────────────────────────────────────────────
    def _find_keyword(self, items, job, existing):
        if existing:
            return existing
        old = job['기존 값']      # 키워드 시트의 "키워드" 칸 (수정·삭제에서는 찾는 값)
        if not old:
            raise PlanError('키워드(또는 키워드 ID)를 적어주세요')
        hit = [k for k in items if _kw_key(k['keyword']) == _kw_key(old)]
        if not hit:
            raise PlanError(f'키워드 "{old}" 이(가) 광고그룹에 없습니다')
        return hit[0]

    def _keyword(self, plan, api, grp, items, job, existing):
        act, gid = job['작업'], grp['nccAdgroupId']
        bid, locked, new = _bid(job['입찰가']), _status(job['상태']), job['새 값']   # new: 등록할 키워드("키워드" 칸)

        def check_new():
            if not new:
                raise PlanError('등록할 키워드를 적어주세요')
            if any(_kw_key(k['keyword']) == _kw_key(new) for k in items):
                raise PlanError(f'키워드 "{new}" 이(가) 이미 있습니다')

        def create_op(bid_, locked_):
            bid_txt = f' 입찰가 {bid_:,}' if bid_ else ' 그룹 입찰가'
            return Op(f'등록 "{new}"({bid_txt.strip()}{", OFF" if locked_ else ""})',
                      lambda: api.create_keyword(gid, new, bid_, locked_))

        if act == '등록':
            check_new()
            plan.ops.append(create_op(bid, locked))
            items.append({'keyword': new, 'nccKeywordId': None})
        elif act == '삭제':
            old = self._find_keyword(items, job, existing)
            plan.ops.append(Op(f'삭제 "{old["keyword"]}"', lambda: api.delete_keyword(old['nccKeywordId']), backup=old))
            self._drop(items, old)
        elif act == '수정':
            old = self._find_keyword(items, job, existing)      # 키워드 문구는 바꿀 수 없다(찾는 값일 뿐): 입찰가·상태만 수정
            if existing and new and _kw_key(new) != _kw_key(old['keyword']):   # ID 로 지정했는데 다른 키워드를 적은 경우
                raise PlanError('키워드 문구는 수정할 수 없습니다. 바꾸려면 등록 줄과 삭제 줄 두 줄로 적어 주세요')
            if bid is None and locked is None:
                raise PlanError('바꿀 값(입찰가·상태)이 없습니다')
            parts = ([f'입찰가 {old.get("bidAmt")}→{bid:,}'] if bid else []) + \
                    ([f'{"OFF" if locked else "ON"}'] if locked is not None else [])
            plan.ops.append(Op(f'수정 "{old["keyword"]}"({", ".join(parts)})',
                               lambda: api.update_keyword(old, bid, locked), backup=old))

    # ── 소재 ─────────────────────────────────────────────────────────
    @staticmethod
    def _require_text_ad(ad):
        """수정은 단일형(TEXT_45) 소재만. 반응형(RSA_AD) 등은 제목·설명이 다른 구조라 건드리면 소재가 망가진다."""
        if ad.get('type') != AD_TYPE:
            raise PlanError(f'{ad_type_label(ad.get("type"))} 소재는 수정할 수 없습니다(단일형 소재만 지원). 삭제는 가능합니다')

    @staticmethod
    def _headline(ad):
        return ((ad.get('ad') or {}).get('headline') or '').strip()

    def _find_ad(self, items, job, existing):
        if existing:
            return existing
        old = job['기존 값']      # 소재 시트의 "제목" 칸 (수정·삭제에서는 찾는 값)
        if not old:
            raise PlanError('제목(또는 소재 ID)을 적어주세요')
        hit = [a for a in items if self._headline(a) == old.strip()]
        if not hit:
            raise PlanError(f'제목이 "{old}" 인 소재가 광고그룹에 없습니다')
        if len(hit) > 1:
            raise PlanError(f'제목이 같은 소재가 {len(hit)}개 있습니다. "대상 ID(선택)" 로 하나를 지정해 주세요')
        return hit[0]

    def _ad(self, plan, api, grp, items, job, existing):
        act, gid = job['작업'], grp['nccAdgroupId']
        locked = _status(job['상태'])
        head, desc, pc, mo = job['새 값'], job['소재 설명'], job['PC URL'], job['모바일 URL']

        def warn_len(h, d):
            if h and '{' not in h and len(h) > config.HEADLINE_MAX:
                plan.warns.append(f'제목 {len(h)}자(권장 {config.HEADLINE_MAX}자 이내)')
            if d and '{' not in d and len(d) > config.DESCRIPTION_MAX:
                plan.warns.append(f'설명 {len(d)}자(권장 {config.DESCRIPTION_MAX}자 이내)')

        def create_op(h, d, p, m, locked_):
            warn_len(h, d)
            return Op(f'등록 "{h}"', lambda: api.create_ad(gid, h, d, p, m, locked_))

        def dup(h, d):
            if any(self._headline(a) == h and (a.get('ad') or {}).get('description') == d for a in items):
                raise PlanError('제목·설명이 같은 소재가 이미 있습니다')

        if act == '등록':
            miss = [n for n, v in (('제목', head), ('소재 설명', desc), ('PC URL', pc)) if not v]
            if miss:
                raise PlanError('등록에는 ' + ', '.join(miss) + ' 이(가) 필요합니다')
            mo = mo or pc          # 모바일 URL 을 비우면 PC URL 과 같게 쓴다
            dup(head, desc)
            plan.ops.append(create_op(head, desc, pc, mo, locked))
            items.append({'ad': {'headline': head, 'description': desc}, 'nccAdId': None})
        elif act == '삭제':
            old = self._find_ad(items, job, existing)
            h = self._headline(old) or f'({ad_type_label(old.get("type"))} 소재, 제목 없음)'
            plan.ops.append(Op(f'삭제 "{h}"', lambda: api.delete_ad(old['nccAdId']), backup=old))
            self._drop(items, old, 'nccAdId')
        elif act == '수정':
            old = self._find_ad(items, job, existing)
            self._require_text_ad(old)      # 제목 칸은 찾는 값일 뿐이라 제목은 바꾸지 않는다: 설명·URL·상태만 수정
            if existing and head and head.strip() != self._headline(old):   # ID 로 지정했는데 다른 제목을 적은 경우
                raise PlanError('제목은 수정할 수 없습니다. 바꾸려면 등록 줄과 삭제 줄 두 줄로 적어 주세요')
            if not (desc or pc or mo) and locked is None:
                raise PlanError('바꿀 값(설명·URL·상태)이 없습니다')
            warn_len('', desc)
            h = self._headline(old)
            plan.ops.append(Op(f'수정 "{h}"', lambda: api.update_ad(old, None, desc, pc, mo, locked), backup=old))

    # ── 공통 ─────────────────────────────────────────────────────────
    @staticmethod
    def _drop(items, old, id_key='nccKeywordId'):
        for i, it in enumerate(items):
            if it.get(id_key) == old.get(id_key):
                del items[i]
                return
