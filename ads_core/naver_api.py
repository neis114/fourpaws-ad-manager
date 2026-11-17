"""네이버 검색광고 API: 파워링크 키워드·소재 조회/등록/수정/삭제."""

import base64
import hashlib
import hmac
import time

import requests

BASE_URL = 'https://api.searchad.naver.com'
POWERLINK_TYPE = 'WEB_SITE'  # 파워링크(사이트 검색 광고) 캠페인 유형
AD_TYPE = 'TEXT_45'          # 파워링크 텍스트 소재(단일형)
AD_TYPE_LABELS = {'TEXT_45': '단일형', 'RSA_AD': '반응형'}   # 화면·안내에 보이는 소재 유형 이름


def ad_type_label(type_):
    """API 소재 유형 → 표시 이름. 모르는 유형은 API 이름 그대로."""
    return AD_TYPE_LABELS.get(type_, type_ or '')


class ApiError(RuntimeError):
    pass


KEEP = object()   # update_extension 의 period 인자: 기간은 바꾸지 않는다
TIMEOUT = (10, 120)   # (연결, 응답 대기) 초. 광고그룹이 크거나 네이버가 느릴 때 30초를 넘기는 경우가 있다
NETWORK_RETRIES = 4   # 조회(GET) 시간 초과·연결 오류 시 시도 횟수


class NaverAdsApi:
    def __init__(self, customer_id, api_key, secret_key):
        self.customer_id = str(customer_id)
        self.api_key = api_key
        self.secret_key = secret_key
        self._campaigns = None

    def _headers(self, method, uri):
        ts = str(int(time.time() * 1000))
        msg = f'{ts}.{method}.{uri}'
        sig = base64.b64encode(
            hmac.new(self.secret_key.encode('utf-8'), msg.encode('utf-8'), hashlib.sha256).digest()
        ).decode('utf-8')
        return {
            'X-Timestamp': ts,
            'X-API-KEY': self.api_key,
            'X-Customer': self.customer_id,
            'X-Signature': sig,
        }

    def _request(self, method, uri, params=None, body=None):
        """요청 1건. 조회(GET)는 시간 초과·연결 오류가 나면 몇 번 다시 시도한다(네이버 응답이 간헐적으로 느리다).
        쓰기(POST/PUT/DELETE)는 서버가 이미 처리했는데 응답만 못 받았을 수 있어 다시 보내지 않는다(중복 등록 방지)."""
        write = method.upper() != 'GET'
        tries = 1 if write else NETWORK_RETRIES
        r, last = None, None
        for attempt in range(tries):
            try:
                for n429 in range(3):   # 요청이 몰려 429 가 오면 잠깐 쉬고 다시 시도 (서명의 시각은 매번 새로 만든다)
                    r = requests.request(method, BASE_URL + uri, headers=self._headers(method, uri),
                                         params=params, json=body, timeout=TIMEOUT)
                    if r.status_code != 429:
                        break
                    time.sleep(1 + n429 * 2)
                break
            except (requests.ConnectionError, requests.Timeout) as e:
                r, last = None, e
                if attempt + 1 < tries:
                    time.sleep(2 * (attempt + 1))
        if r is None:
            if write:
                raise ApiError(f'API {method} {uri} 응답을 받지 못했습니다(시간 초과·연결 오류). 서버가 요청을 이미 처리했을 수 있으니 '
                               '광고주센터에서 반영 여부를 확인한 뒤 필요하면 다시 실행하세요')
            raise ApiError(f'API {method} {uri} 응답이 없습니다(시간 초과·연결 오류, {tries}회 시도): {type(last).__name__}')
        if not r.ok:
            raise ApiError(f'API {method} {uri} 실패 ({r.status_code}): {r.text[:300]}')
        return r.json() if r.text else None

    # ── 조회 ─────────────────────────────────────────────────────────
    def powerlink_campaigns(self):
        if self._campaigns is None:
            self._campaigns = [c for c in self._request('GET', '/ncc/campaigns') if c.get('campaignTp') == POWERLINK_TYPE]
        return self._campaigns

    def campaign_summary(self):
        """계정의 캠페인 유형 분포(안내용). 예: "PLACE 1, BRAND_SEARCH 5"."""
        counts = {}
        for c in self._request('GET', '/ncc/campaigns') or []:
            counts[c.get('campaignTp')] = counts.get(c.get('campaignTp'), 0) + 1
        return ', '.join(f'{k} {v}' for k, v in counts.items()) or '캠페인 없음'

    def find_adgroup(self, campaign, adgroup):
        """캠페인·광고그룹 이름(앞뒤 공백 무시, 정확히 일치)으로 광고그룹 1개를 찾는다. 못 찾거나 여럿이면 LookupError."""
        camps = [c for c in self.powerlink_campaigns() if c.get('name', '').strip() == campaign.strip()]
        if not camps:
            raise LookupError(f'파워링크 캠페인 "{campaign}" 을(를) 찾지 못했습니다')
        groups = []
        for camp in camps:
            groups += [g for g in self._request('GET', '/ncc/adgroups', {'nccCampaignId': camp['nccCampaignId']})
                       if g.get('name', '').strip() == adgroup.strip()]
        if not groups:
            raise LookupError(f'캠페인 "{campaign}" 에서 광고그룹 "{adgroup}" 을(를) 찾지 못했습니다')
        if len(groups) > 1:
            raise LookupError(f'광고그룹 "{adgroup}" 이(가) {len(groups)}개 있어 하나로 정할 수 없습니다')
        return groups[0]

    def adgroups(self, campaign_id):
        return self._request('GET', '/ncc/adgroups', {'nccCampaignId': campaign_id}) or []

    def get_adgroup(self, adgroup_id):
        return self._request('GET', f'/ncc/adgroups/{adgroup_id}')

    def get_keyword(self, keyword_id):
        return self._request('GET', f'/ncc/keywords/{keyword_id}')

    def get_ad(self, ad_id):
        return self._request('GET', f'/ncc/ads/{ad_id}')

    def keywords(self, adgroup_id):
        return self._request('GET', '/ncc/keywords', {'nccAdgroupId': adgroup_id}) or []

    def ads(self, adgroup_id):
        return self._request('GET', '/ncc/ads', {'nccAdgroupId': adgroup_id}) or []

    # ── 키워드 ───────────────────────────────────────────────────────
    def create_keyword(self, adgroup_id, keyword, bid=None, locked=False):
        """키워드 1개 등록. bid 가 없으면 광고그룹 입찰가를 따른다."""
        item = {'keyword': keyword, 'userLock': bool(locked)}
        if bid is None:
            item['useGroupBidAmt'] = True
        else:
            item.update({'useGroupBidAmt': False, 'bidAmt': int(bid)})
        res = self._request('POST', '/ncc/keywords', {'nccAdgroupId': adgroup_id}, [item])
        if not res:
            raise ApiError('키워드 등록 응답이 비어 있습니다')
        return res[0]

    def update_keyword(self, kw, bid=None, locked=None):
        """키워드의 입찰가·ON/OFF 수정. 바꿀 값만 넘긴다."""
        fields, body = [], {'nccKeywordId': kw['nccKeywordId'], 'nccAdgroupId': kw['nccAdgroupId']}
        if bid is not None:
            fields += ['bidAmt', 'useGroupBidAmt']
            body.update({'bidAmt': int(bid), 'useGroupBidAmt': False})
        if locked is not None:
            fields.append('userLock')
            body['userLock'] = bool(locked)
        if not fields:
            return kw
        return self._request('PUT', f'/ncc/keywords/{kw["nccKeywordId"]}', {'fields': ','.join(fields)}, body)

    def delete_keyword(self, keyword_id):
        self._request('DELETE', f'/ncc/keywords/{keyword_id}')

    # ── 소재 ─────────────────────────────────────────────────────────
    def create_ad(self, adgroup_id, headline, description, pc_url, mo_url, locked=False):
        body = {
            'nccAdgroupId': adgroup_id, 'type': AD_TYPE, 'userLock': bool(locked),
            'ad': {'headline': headline, 'description': description,
                   'pc': {'final': pc_url}, 'mobile': {'final': mo_url}},
        }
        return self._request('POST', '/ncc/ads', {}, body)

    def update_ad(self, ad, headline=None, description=None, pc_url=None, mo_url=None, locked=None):
        """소재의 문구·URL·ON/OFF 수정. 바꿀 값만 넘긴다 (문구·URL 과 ON/OFF 는 요청을 나눠 보낸다)."""
        res = ad
        content = {k: v for k, v in (('headline', headline), ('description', description)) if v}
        if content or pc_url or mo_url:
            # 표시용 URL(display)·punyCode 는 네이버가 최종 URL 에서 다시 만드므로 보내지 않는다
            cur = ad.get('ad') or {}
            new_ad = {
                'headline': cur.get('headline'), 'description': cur.get('description'),
                'pc': {'final': (cur.get('pc') or {}).get('final')},
                'mobile': {'final': (cur.get('mobile') or {}).get('final')},
            }
            new_ad.update(content)
            if pc_url:
                new_ad['pc'] = {'final': pc_url}
            if mo_url:
                new_ad['mobile'] = {'final': mo_url}
            res = self._request('PUT', f'/ncc/ads/{ad["nccAdId"]}', {'fields': 'ad'},
                                {'nccAdId': ad['nccAdId'], 'nccAdgroupId': ad['nccAdgroupId'], 'ad': new_ad})
        if locked is not None:
            res = self._request('PUT', f'/ncc/ads/{ad["nccAdId"]}', {'fields': 'userLock'},
                                {'nccAdId': ad['nccAdId'], 'nccAdgroupId': ad['nccAdgroupId'], 'userLock': bool(locked)})
        return res

    def delete_ad(self, ad_id):
        self._request('DELETE', f'/ncc/ads/{ad_id}')

    # ── 확장소재 (추가제목·추가설명·홍보문구·서브링크·이미지형 서브링크·파워링크 이미지) ─────────────
    def extensions(self, adgroup_id):
        return self._request('GET', '/ncc/ad-extensions', {'ownerId': adgroup_id}) or []

    def get_extension(self, ext_id):
        return self._request('GET', f'/ncc/ad-extensions/{ext_id}')

    def create_extension(self, grp, type_, ad_extension, locked=False, period=None):
        """확장소재 1개 등록. period 는 None 또는 (시작 ISO, 종료 ISO). 연결 채널은 광고그룹의 것을 따른다."""
        body = {'ownerId': grp['nccAdgroupId'], 'type': type_, 'adExtension': ad_extension,
                'userLock': bool(locked), 'usePeriod': bool(period)}
        for key in ('pcChannelId', 'mobileChannelId'):
            if grp.get(key):
                body[key] = grp[key]
        if period:
            body['periodStartDt'], body['periodEndDt'] = period
        return self._request('POST', '/ncc/ad-extensions', {}, body)

    def update_extension(self, ext, ad_extension=None, locked=None, period=KEEP):
        """확장소재의 내용·ON/OFF·기간 수정. 바꿀 것만 넘긴다(항목마다 요청을 나눠 보낸다). period 는 None 이면 기간 해제."""
        def put(extra):
            body = {'nccAdExtensionId': ext['nccAdExtensionId'], 'ownerId': ext['ownerId'], 'type': ext['type'], **extra}
            return self._request('PUT', f'/ncc/ad-extensions/{ext["nccAdExtensionId"]}', {'fields': ','.join(extra)}, body)

        res = ext
        if ad_extension is not None:
            res = put({'adExtension': ad_extension})
        if period is not KEEP:
            res = put({'usePeriod': False} if period is None else
                      {'usePeriod': True, 'periodStartDt': period[0], 'periodEndDt': period[1]})
        if locked is not None:
            res = put({'userLock': bool(locked)})
        return res

    def delete_extension(self, ext_id):
        self._request('DELETE', f'/ncc/ad-extensions/{ext_id}')
