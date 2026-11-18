"""미리보기 / 실행 / 현황 가져오기. input 의 작업 엑셀을 읽고, 계획을 세우고, 결과를 output 의 사본 엑셀에 기록한다."""

import collections
import datetime as dt
import glob
import json
import os
import time

from . import config, specs
from .excel_io import ResultWriter, read_jobs, write_snapshot_file
from .naver_api import ApiError, ad_type_label
from .planner import Planner

FLUSH_EVERY = 10          # 이 개수마다 결과를 파일에 먼저 저장한다(중간에 멈춰도 처리된 행이 "완료"로 남게)
KEYWORD_HEADER = ['계정', '캠페인', '광고그룹', '키워드', '입찰가', '입찰가 방식', '상태', '검수상태',
                  '광고그룹 ID', '키워드 ID']
AD_HEADER = ['계정', '캠페인', '광고그룹', '제목', '소재 설명', 'PC URL', '모바일 URL', '상태', '검수상태', '소재 유형',
             '광고그룹 ID', '소재 ID']     # 소재 유형: 단일형(TEXT_45) / 반응형(RSA_AD, 제목·설명이 소재 구성요소에 있어 비어 보인다)
CAMPAIGN_HEADER = ['계정', '캠페인', '캠페인 ID', '캠페인 유형', '하루 예산', '상태', '검수상태']
ADGROUP_HEADER = ['계정', '캠페인', '광고그룹', '광고그룹 ID', '기본 입찰가', '상태', '검수상태']


def _tag(p):
    """로그·안내용 행 표시: 키워드 3행 / 소재 5행"""
    return f'{p.kind} {p.row}행'


def _result(p, text, time_=''):
    return {'sheet': p.kind, 'row': p.row, 'result': text, 'time': time_}


def _now():
    return dt.datetime.now().strftime('%Y-%m-%d %H:%M')


def _stem(path):
    return os.path.splitext(os.path.basename(path))[0]


def _check_input(path):
    if not path or not os.path.isfile(path):
        raise FileNotFoundError(f'작업 엑셀을 찾지 못했습니다: {path or "(선택 안 됨)"}  (input 폴더에 .xlsx 를 올려 주세요)')


def _load_jobs(path, log):
    """작업 엑셀의 모든 시트(키워드·소재·확장소재)에서 처리 대상 행을 모은다. 이미 "완료"인 행과 계정·작업이 모두 빈 행은 건너뛴다."""
    _check_input(path)
    try:
        data = read_jobs(path)
    except Exception as e:  # noqa: BLE001
        raise RuntimeError(f'엑셀을 읽지 못했습니다: {e}')
    if not data:
        raise RuntimeError('작업 시트(키워드·소재·확장소재)를 하나도 찾지 못했습니다. 템플릿 엑셀을 복사해 쓴 파일인지 확인하세요.')
    jobs, done = [], 0
    for name, block in data.items():
        spec = specs.BY_SHEET[name]
        missing = [c for c in spec['cols'] if c not in block['headers']]
        if missing:
            raise RuntimeError(f'"{name}" 시트 1행에 없는 열: {", ".join(missing)}')
        for rec in block['rows']:
            raw = {h: ('' if v is None else str(v)).strip() for h, v in rec['values'].items()}
            if not raw.get('작업') and not raw.get('계정'):
                continue
            job = {c: '' for c in config.COLS}          # planner 가 쓰는 내부 항목 이름(시트별 열 이름을 연결)
            job.update(raw)
            for internal, header in spec['map'].items():
                job[internal] = raw.get(header, '')
            job['대상'], job['row'] = name, rec['row']
            if job['결과'].startswith('완료'):
                done += 1
                continue
            jobs.append(job)
    kinds = ' · '.join(f'{k} {n}' for k in specs.KINDS if (n := sum(1 for j in jobs if j['대상'] == k)))
    log(f'작업 엑셀: 처리할 행 {len(jobs)}개 ({kinds or "없음"}) / 이미 완료된 {done}개는 건너뜀')
    jobs, moved = _delete_first(jobs)
    if moved:
        log(f'같은 항목의 삭제와 등록이 함께 있어 삭제 {moved}건을 등록보다 먼저 처리합니다.')
    return jobs


def _identity(job):
    """같은 항목을 가리키는 키: (대상, 계정, 광고그룹, 값). 값 없이 ID 로만 지정한 행은 None."""
    kind = job['대상']
    if kind == '키워드':
        text = ''.join(job['기존 값'].split()).lower()      # 키워드는 띄어쓰기·대소문자만 달라도 같은 것으로 본다
    elif kind == '소재':
        text = job['기존 값'].strip()
    else:
        text = job.get(specs.EXT[kind].key_col, '').strip()
    if not text:
        return None
    group = job['광고그룹 ID(선택)'] or (job['캠페인'].strip(), job['광고그룹'].strip())
    return kind, job['계정'].strip().upper(), group, text


def _delete_first(jobs):
    """같은 항목의 삭제와 등록이 함께 있으면 삭제를 먼저 처리한다(같은 값을 다시 등록하려면 기존 것이 먼저 지워져야 한다).
    그 밖의 순서는 엑셀에 적은 그대로다. 반환: (순서를 맞춘 행 목록, 등록보다 앞으로 당겨진 삭제 수)"""
    pos = {id(j): i for i, j in enumerate(jobs)}
    deletes = collections.defaultdict(list)
    for j in jobs:
        if j['작업'] == '삭제' and (key := _identity(j)):
            deletes[key].append(j)
    out, emitted, moved = [], set(), 0
    for j in jobs:
        if id(j) in emitted:
            continue
        if j['작업'] == '등록' and (key := _identity(j)):
            for d in deletes.get(key, ()):
                if id(d) not in emitted:
                    out.append(d)
                    emitted.add(id(d))
                    moved += pos[id(d)] > pos[id(j)]
        out.append(j)
        emitted.add(id(j))
    return out, moved


def _counts(plans):
    ok = [p for p in plans if not p.error]
    c = {a: sum(1 for p in ok if p.action == a) for a in config.ACTIONS}
    c['오류'] = len(plans) - len(ok)
    return c


def _fmt_counts(c):
    return ' · '.join(f'{k} {v}' for k, v in c.items() if v) or '없음'


def preview(path, env=None, log=print):
    """실제 변경 없이 행마다 무엇을 할지 계산하고, output 의 "<이름>_미리보기.xlsx" 결과 칸에 "예정/오류"를 적는다."""
    env = env or config.load_env()
    log(f'.env: {env["_PATH"]}')
    log(f'입력 파일: {path}')
    jobs = _load_jobs(path, log)
    planner = Planner(env)
    plans = []
    for job in jobs:
        plan = planner.plan(job)
        plans.append(plan)
        log(f'  [{_tag(plan)}] {plan.preview()}')
    writer = ResultWriter(path, os.path.join(config.OUTPUT_DIR, f'{_stem(path)}_미리보기.xlsx'))
    writer.add([_result(p, p.preview()) for p in plans])
    out = writer.save()
    log(f'\n미리보기 완료: {_fmt_counts(_counts(plans))}  (실제 변경은 하지 않았습니다)')
    log(f'결과 파일: {out}')
    return plans


def _backup(plans):
    """바뀌거나 사라지는 기존 항목을 JSON 으로 남긴다(되돌릴 때 참고). 파일 경로를 돌려준다."""
    items = [{'row': p.row, 'account': p.account, 'op': op.desc, 'before': op.backup}
             for p in plans for op in p.ops if op.backup]
    if not items:
        return None
    os.makedirs(config.BACKUP_DIR, exist_ok=True)
    path = os.path.join(config.BACKUP_DIR, dt.datetime.now().strftime('%Y%m%d_%H%M%S') + '.json')
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(items, f, ensure_ascii=False, indent=2)
    return path


def apply(path, env=None, log=print, confirm=None, should_stop=None):
    """계획을 다시 세운 뒤 실제로 실행하고, 결과를 output 의 "<이름>_결과_<일시>.xlsx" 에 기록한다.

    :param confirm: 함수(요약 문자열) -> bool. 실행 직전에 호출하며 False 면 아무것도 바꾸지 않고 끝낸다.
    :param should_stop: True 를 돌려주면 현재 행까지만 처리하고 멈춘다.
    """
    env = env or config.load_env()
    log(f'.env: {env["_PATH"]}')
    log(f'입력 파일: {path}')
    jobs = _load_jobs(path, log)
    planner = Planner(env)
    plans = []
    for job in jobs:
        plan = planner.plan(job)
        plans.append(plan)
        log(f'  [{_tag(plan)}] {plan.preview()}')
    runnable = [p for p in plans if not p.error]
    summary = _fmt_counts(_counts(plans))
    stamp = f'{dt.datetime.now():%Y%m%d_%H%M%S}'
    errors = [p for p in plans if p.error]
    error_rows = [_result(p, p.preview()) for p in errors]
    if not runnable:    # 실행할 게 없으면 오류 사유만 담은 파일을 남긴다(이전 "실행"으로 세지 않도록 이름을 구분)
        if errors:
            ew = ResultWriter(path, os.path.join(config.OUTPUT_DIR, f'{_stem(path)}_오류_{stamp}.xlsx'))
            ew.add(error_rows)
            log(f'오류 사유 파일: {ew.save()}')
        log(f'\n실행할 행이 없습니다. ({summary})')
        return {'done': 0, 'failed': 0, 'errors': len(errors)}
    earlier = glob.glob(os.path.join(config.OUTPUT_DIR, f'{_stem(path)}_결과_*.xlsx'))
    if earlier:   # 같은 파일을 다시 올리면 "완료" 표시가 없어 이미 처리한 행이 중복 실행될 수 있다
        summary += ('\n\n⚠️ 이 파일은 이미 실행한 결과 파일이 output 에 있습니다. 같은 파일을 다시 실행하면 처리한 행이 중복될 수 있으니, '
                    '일부만 다시 하려면 output 의 결과 파일(완료 표시됨)을 input 에 올려 실행하세요.')
    if confirm and not confirm(summary):
        log('\n취소했습니다. 아무것도 바꾸지 않았습니다.')
        return {'done': 0, 'failed': 0, 'errors': len(errors), 'cancelled': True}
    writer = ResultWriter(path, os.path.join(config.OUTPUT_DIR, f'{_stem(path)}_결과_{stamp}.xlsx'))
    writer.add(error_rows)      # 실행 전에 걸러진 오류 행의 사유도 결과 파일에 함께 적는다

    backup_path = _backup(runnable)
    if backup_path:
        log(f'백업: {backup_path}')

    results, done, failed = [], 0, 0

    def flush():
        if results:
            writer.add(list(results))
            writer.save()
            results.clear()

    try:
        for p in runnable:
            if should_stop and should_stop():
                log('⏹ 중지 요청으로 여기까지만 처리합니다.')
                break
            msgs, ok = [], True
            for op in p.ops:
                try:
                    op.run()
                    msgs.append(op.desc)
                    time.sleep(0.1)
                except Exception as e:  # noqa: BLE001
                    ok = False
                    msgs.append(f'{op.desc} 실패({e})')
                    break      # 한 행의 작업이 실패하면 그 행의 나머지 작업은 하지 않는다
            if ok:
                done += 1
                text = '완료: ' + ' → '.join(msgs)
            else:
                failed += 1
                text = '실패: ' + ' → '.join(msgs)
            log(f'  [{_tag(p)}] {text}')
            results.append(_result(p, text, _now()))
            if len(results) >= FLUSH_EVERY:
                flush()
    finally:
        flush()
    log(f'\n실행 완료: 성공 {done}행 / 실패 {failed}행 / 사전 오류 {len(errors)}행')
    log(f'결과 파일: {writer.dst}  (실패·오류 행은 고쳐서 이 파일을 input 에 올리면 완료된 행은 건너뛰고 다시 실행합니다)')
    return {'done': done, 'failed': failed, 'errors': len(errors), 'backup': backup_path, 'result_file': writer.dst}


def snapshot(accounts, env=None, log=print, targets=None):
    """계정의 파워링크 현황을 output 의 "현황_<일시>.xlsx" 로 저장한다.
    targets: 수집 대상 리스트 (예: ['전체'], ['캠페인', '그룹'], ['키워드'], ['소재'], ['확장소재'])
    미지정 또는 '전체' 포함 시 모든 시트(캠페인·그룹·키워드·소재·확장소재 6종)를 수집한다.
    """
    env = env or config.load_env()
    planner = Planner(env)

    # 타겟 파싱
    if not targets:
        inc_camp = inc_grp = inc_kw = inc_ad = inc_ext = True
    else:
        is_all = '전체' in targets
        inc_camp = is_all or ('캠페인' in targets)
        inc_grp = is_all or ('그룹' in targets)
        inc_kw = is_all or ('키워드' in targets)
        inc_ad = is_all or ('소재' in targets)
        inc_ext = is_all or ('확장소재' in targets)

    target_labels = []
    if inc_camp: target_labels.append('캠페인')
    if inc_grp: target_labels.append('그룹')
    if inc_kw: target_labels.append('키워드')
    if inc_ad: target_labels.append('소재')
    if inc_ext: target_labels.append('확장소재')
    log(f'수집 대상 항목: {", ".join(target_labels)}')

    camp_rows = [CAMPAIGN_HEADER] if inc_camp else []
    grp_rows = [ADGROUP_HEADER] if inc_grp else []
    kw_rows = [KEYWORD_HEADER] if inc_kw else []
    ad_rows = [AD_HEADER] if inc_ad else []
    ext_rows = {k.name: [k.snapshot_header] for k in specs.EXT_KINDS} if inc_ext else {}
    ext_by_type = {k.type: k for k in specs.EXT_KINDS} if inc_ext else {}

    failed = []       # 조회에 끝내 실패한 광고그룹
    need_groups = inc_grp or inc_kw or inc_ad or inc_ext

    for code in accounts:
        api = planner.api(code)
        log(f'{code}: 파워링크 현황 조회 중...')
        camps = api.powerlink_campaigns()
        if not camps:
            log(f'  ⚠️ {code}: 파워링크(WEB_SITE) 캠페인이 없어 건너뜁니다. 이 계정의 캠페인: {api.campaign_summary()} '
                '(이 도구는 파워링크만 다룹니다)')
            continue

        for camp in camps:
            cid = camp['nccCampaignId']
            cname = camp['name']

            if inc_camp:
                camp_rows.append([
                    code.upper(), cname, cid,
                    camp.get('campaignType', 'WEB_SITE'),
                    camp.get('dailyBudget', ''),
                    'OFF' if camp.get('userLock') else 'ON',
                    camp.get('inspectStatus', '')
                ])

            if not need_groups:
                continue

            groups = api.adgroups(cid)
            log(f'  · {cname} ({len(groups)}개 광고그룹)')

            for grp in groups:
                gid = grp['nccAdgroupId']
                gname = grp['name']
                base = [code.upper(), cname, gname]

                if inc_grp:
                    grp_rows.append([
                        code.upper(), cname, gname, gid,
                        grp.get('bidAmt', ''),
                        'OFF' if grp.get('userLock') else 'ON',
                        grp.get('inspectStatus', '')
                    ])

                # 키워드/소재/확장소재가 모두 불필요하면 광고그룹 하위 API 호출은 생략 (초고속!)
                if not (inc_kw or inc_ad or inc_ext):
                    continue

                keywords, ads, exts = [], [], []
                try:
                    if inc_kw:
                        keywords = api.keywords(gid)
                    if inc_ad:
                        ads = api.ads(gid)
                    if inc_ext:
                        exts = api.extensions(gid)
                except ApiError as e:
                    failed.append(f'{code} / {cname} / {gname}')
                    log(f'    ⚠️ 조회 실패로 건너뜀: {gname} ({e})')
                    continue

                if inc_kw:
                    for k in keywords:
                        own = not k.get('useGroupBidAmt')
                        kw_rows.append(base + [k['keyword'], k.get('bidAmt') if own else grp.get('bidAmt'),
                                               '개별' if own else '그룹', 'OFF' if k.get('userLock') else 'ON',
                                               k.get('inspectStatus', ''), gid, k['nccKeywordId']])
                if inc_ad:
                    for a in ads:
                        ad = a.get('ad') or {}
                        ad_rows.append(base + [ad.get('headline', ''), ad.get('description', ''),
                                               (ad.get('pc') or {}).get('final', ''),
                                               (ad.get('mobile') or {}).get('final', ''),
                                               'OFF' if a.get('userLock') else 'ON', a.get('inspectStatus', ''),
                                               ad_type_label(a.get('type')), gid, a['nccAdId']])
                if inc_ext:
                    for x in exts:
                        kind = ext_by_type.get(x.get('type'))
                        if not kind or not x.get('adExtension'):
                            continue
                        cur = kind.read(x['adExtension'])
                        ext_rows[kind.name].append(
                            base + [cur[c] for c in kind.values]
                            + [specs.iso_to_date(x.get('periodStartDt')) if x.get('usePeriod') else '',
                               specs.iso_to_date(x.get('periodEndDt')) if x.get('usePeriod') else '',
                               'OFF' if x.get('userLock') else 'ON', x.get('inspectStatus', ''), gid, x['nccAdExtensionId']])

        # 계정별 누적 로그
        counts_msg = []
        if inc_camp: counts_msg.append(f'캠페인 {len(camp_rows) - 1}행')
        if inc_grp: counts_msg.append(f'그룹 {len(grp_rows) - 1}행')
        if inc_kw: counts_msg.append(f'키워드 {len(kw_rows) - 1}행')
        if inc_ad: counts_msg.append(f'소재 {len(ad_rows) - 1}행')
        if inc_ext:
            counts_msg.extend(f'{n} {len(r) - 1}행' for n, r in ext_rows.items())
        log(f'  {code}: {" · ".join(counts_msg)} (누적)')

    tables = {}
    if inc_camp and len(camp_rows) > 1:
        tables['캠페인'] = camp_rows
    if inc_grp and len(grp_rows) > 1:
        tables['그룹'] = grp_rows
    if inc_kw and len(kw_rows) > 1:
        tables['키워드'] = kw_rows
    if inc_ad and len(ad_rows) > 1:
        tables['소재'] = ad_rows
    if inc_ext:
        for k_name, r in ext_rows.items():
            if len(r) > 1:
                tables[k_name] = r

    if not tables:
        log('\n가져온 현황이 없어 파일을 만들지 않았습니다. 선택한 대상 항목과 계정을 확인하세요.')
        return None

    out = write_snapshot_file(os.path.join(config.OUTPUT_DIR, f'현황_{dt.datetime.now():%Y%m%d_%H%M%S}.xlsx'), tables)
    total = sum(len(r) - 1 for r in tables.values())
    log(f'\n현황 {total}행을 시트 {len(tables)}개로 저장했습니다: {out}')
    if failed:
        log(f'\n⚠️ 조회에 실패한 광고그룹 {len(failed)}개는 이 파일에서 빠져 있습니다. 네이버 응답이 느린 시간대였을 수 있으니 잠시 뒤 다시 받아 보세요.')
        for name in failed[:15]:
            log(f'   - {name}')
        if len(failed) > 15:
            log(f'   … 외 {len(failed) - 15}개')
    return out
