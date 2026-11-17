"""확장소재(추가제목·추가설명·홍보문구·서브링크·이미지형 서브링크·파워링크 이미지) 한 행 → 실행 계획.

키워드·소재와 같은 규칙: 시트의 기본 값 열(추가제목·링크1 이름·이미지 ID …)이 등록에서는 새 값, 수정·삭제에서는 찾는 값이다.
확장소재 ID 칸을 적으면 ID가 우선. 수정에서 비워 둔 값 칸은 기존 값을 유지한다. 값 자체를 바꾸는 교체는 없다(등록 줄 + 삭제 줄).
"""

import copy

from . import specs
from .naver_api import KEEP
from .planner import Op, PlanError, _status

IMAGE_TYPES = ('POWER_LINK_IMAGE', 'IMAGE_SUB_LINKS')


class ImageContext:
    """이미지 ID(또는 imagePath 일부) → imagePath. 이미지 목록을 조회하는 API 가 없어,
    같은 계정의 기존 이미지형 확장소재가 쓰는 이미지에서 찾는다(계정당 한 번만 훑는다)."""

    def __init__(self, planner, api):
        self.planner, self.api = planner, api

    def image(self, text):
        index = self.planner.image_index(self.api)
        t = text.strip()
        hit = {p: v for p, v in index.items() if t == p or t in p}
        if not hit:
            raise specs.SpecError(
                f'이미지 "{t}" 을(를) 이 계정의 기존 확장소재에서 찾지 못했습니다. '
                '현황 엑셀의 파워링크이미지 시트에서 이미지 ID(전체 경로)를 복사해 적거나, 이미지를 광고주센터에서 먼저 한 번 사용해 주세요')
        if len(hit) > 1:
            raise specs.SpecError(f'이미지 "{t}" 에 해당하는 이미지가 {len(hit)}개입니다. 더 긴 값을 적어주세요')
        return next(iter(hit.values()))


def _period(job):
    """기간 칸 → KEEP(비움) / None("해제") / (시작 ISO, 종료 ISO)."""
    start, end = job['기간 시작일'].strip(), job['기간 종료일'].strip()
    if not start and not end:
        return KEEP
    if start in ('해제', '없음') or end in ('해제', '없음'):
        return None
    if not (start and end):
        raise PlanError('기간은 시작일과 종료일을 모두 적어주세요 (기간을 없애려면 시작일에 "해제")')
    try:
        s, e = specs.date_to_iso(start), specs.date_to_iso(end)
    except specs.SpecError as err:
        raise PlanError(str(err))
    if s > e:
        raise PlanError('기간 종료일이 시작일보다 앞입니다')
    return s, e


def plan_extension(planner, plan, api, grp, items, job, existing, kind):
    spec = specs.EXT[kind]
    act = job['작업']
    mine = [x for x in items if x.get('type') == spec.type]
    vals = {c: job.get(c, '').strip() for c in spec.values}
    key = vals[spec.key_col]
    # 수정에서 기본 값 열은 찾는 값일 뿐이라, 나머지 값 열에 적은 것만 "바꿀 값"이다
    has_new = any(v for c, v in vals.items() if c != spec.key_col)
    locked = _status(job['상태'])
    period = _period(job)
    ctx = ImageContext(planner, api)

    def build(v):
        try:
            return spec.build(v, ctx)
        except specs.SpecError as e:
            raise PlanError(str(e))

    def find_old():
        if existing:
            if existing.get('type') != spec.type:
                raise PlanError(f'확장소재 ID 의 유형이 다릅니다({existing.get("type")}). "{kind}" 시트에는 {spec.type} 만 쓸 수 있습니다')
            return existing
        if not key:
            raise PlanError(f'"{spec.key_col}"(또는 확장소재 ID)을 적어주세요')
        hit = [x for x in mine if spec.match(key, x.get('adExtension'))]
        if not hit:
            raise PlanError(f'{kind} "{key}" 이(가) 광고그룹에 없습니다')
        if len(hit) > 1:
            raise PlanError(f'"{key}" 에 해당하는 {kind}이(가) {len(hit)}개 있습니다. "확장소재 ID(선택)" 로 하나를 지정해 주세요')
        return hit[0]

    def same(ad, others):
        sig = spec.read(ad)
        return any(spec.read(x.get('adExtension') or {}) == sig for x in others)

    def merged(old):
        cur = spec.read(old.get('adExtension') or {})
        cur.update({k: v for k, v in vals.items() if v})
        return cur

    def drop(old):
        for i, it in enumerate(items):
            if it.get('nccAdExtensionId') == old.get('nccAdExtensionId'):
                del items[i]
                return

    def create_op(ad, locked_, period_):
        per = f' 기간 {specs.iso_to_date(period_[0])}~{specs.iso_to_date(period_[1])}' if period_ else ''
        return Op(f'등록 "{spec.label(ad)}"{per}{" (OFF)" if locked_ else ""}',
                  lambda: api.create_extension(grp, spec.type, ad, locked_, period_))

    def delete_op(old):
        return Op(f'삭제 "{spec.label(old.get("adExtension") or {})}"',
                  lambda: api.delete_extension(old['nccAdExtensionId']), backup=old)

    if act == '등록':
        ad = build(vals)
        if same(ad, mine):
            raise PlanError('같은 내용의 확장소재가 이미 있습니다')
        plan.ops.append(create_op(ad, locked, None if period is KEEP else period))
        items.append({'type': spec.type, 'adExtension': ad, 'nccAdExtensionId': None})

    elif act == '삭제':
        old = find_old()
        plan.ops.append(delete_op(old))
        drop(old)

    else:  # 수정
        old = find_old()
        old_ad = old.get('adExtension') or {}
        if existing and key and not spec.match(key, old_ad):    # ID 로 지정했는데 기본 값 칸에 다른 값을 적은 경우: 조용히 무시하지 않는다
            raise PlanError(f'"{spec.key_col}" 칸의 값은 수정할 수 없습니다. 값을 바꾸려면 등록 줄과 삭제 줄 두 줄로 적어 주세요')
        new_ad = None
        if has_new:
            cand = build(merged(old))
            if spec.read(cand) != spec.read(old_ad):
                if same(cand, [x for x in mine if x is not old]):
                    raise PlanError('같은 내용의 확장소재가 이미 있습니다')
                new_ad = cand
        if new_ad is None and locked is None and period is KEEP:
            raise PlanError('바꿀 값이 없거나 기존과 같습니다')
        parts = (['내용'] if new_ad is not None else []) + (['기간'] if period is not KEEP else []) + \
                ([('OFF' if locked else 'ON')] if locked is not None else [])
        plan.ops.append(Op(f'수정 "{spec.label(old_ad)}"({", ".join(parts)})',
                           lambda: api.update_extension(old, new_ad, locked, period), backup=copy.deepcopy(old)))
        if new_ad is not None:
            old['adExtension'] = new_ad
