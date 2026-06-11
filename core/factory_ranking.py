"""
factory_ranking.py — 공장·지역 불량률 분석 및 랭킹 계산 모듈
"""
from __future__ import annotations
from collections import defaultdict
from datetime import datetime


def _parse_date(val) -> str | None:
    """날짜 문자열을 YYYY-MM 형식으로 정규화 (regex 기반)"""
    if not val:
        return None
    if hasattr(val, 'strftime'):
        return val.strftime('%Y-%m')
    if isinstance(val, (int, float)) and 30000 < val < 60000:
        try:
            from datetime import timedelta
            return (datetime(1899, 12, 30) + timedelta(days=int(val))).strftime('%Y-%m')
        except Exception:
            pass
    import re
    s = str(val).strip()
    m = re.match(r'(\d{4})[-./](\d{1,2})', s)
    if m:
        y, mo = int(m.group(1)), int(m.group(2))
        if 2000 <= y <= 2099 and 1 <= mo <= 12:
            return f'{y}-{mo:02d}'
    m = re.match(r'(\d{4})(\d{2})\d{2}$', s)
    if m:
        y, mo = int(m.group(1)), int(m.group(2))
        if 2000 <= y <= 2099 and 1 <= mo <= 12:
            return f'{y}-{mo:02d}'
    return None


def _safe_rate(qty_total: int, inspec: int) -> float:
    """불량률(%) — 검사수량 0이면 None"""
    if not inspec:
        return None
    return round(qty_total / inspec * 100, 2)


# ── 공장별 집계 ──────────────────────────────────────────────────
def calc_factory_ranking(raw_rows: list[dict], cache: dict,
                         start: str = None, end: str = None,
                         buyer: str = None, item: str = None) -> list[dict]:
    """
    공장별 불량률·순위 계산
    반환: [{ factory, region1, region2, avg_rate, total_inspec,
             total_defect, record_count, trend, monthly }]
    trend: 'down' | 'up' | 'flat' | 'new'  (최근 3개월 기준)
    """
    # 필터링
    rows = _filter_rows(raw_rows, start, end, buyer, item)

    # 공장별 월별 집계
    factory_monthly: dict[str, dict[str, dict]] = defaultdict(lambda: defaultdict(
        lambda: {'inspec': 0, 'defect': 0, 'count': 0}
    ))
    factory_info: dict[str, dict] = {}

    for r in rows:
        f = r.get('factory') or '(미입력)'
        if not f or f == '(미입력)':
            continue
        ym = _parse_date(r.get('date')) or 'unknown'
        factory_monthly[f][ym]['inspec'] += r.get('inspec', 0) or 0
        factory_monthly[f][ym]['defect'] += r.get('qty_total', 0) or 0
        factory_monthly[f][ym]['count']  += 1
        if f not in factory_info:
            factory_info[f] = {
                'region1':       r.get('region1', ''),
                'region2':       r.get('region2', ''),
                'region_label':  r.get('region_label', r.get('region1', '')),
            }

    result = []
    for factory, monthly_data in factory_monthly.items():
        months = sorted(m for m in monthly_data if m != 'unknown')
        total_inspec = sum(v['inspec'] for v in monthly_data.values())
        total_defect = sum(v['defect'] for v in monthly_data.values())
        total_count  = sum(v['count']  for v in monthly_data.values())

        avg_rate = _safe_rate(total_defect, total_inspec)

        # 월별 불량률 시계열
        monthly_rates = []
        for m in months:
            d = monthly_data[m]
            monthly_rates.append({
                'month': m,
                'rate': _safe_rate(d['defect'], d['inspec']),
                'inspec': d['inspec'],
                'defect': d['defect'],
            })

        # 추이: 최근 3개월 선형 기울기
        trend = _calc_trend(monthly_rates)

        result.append({
            'factory':      factory,
            'region1':      factory_info[factory]['region1'],
            'region2':      factory_info[factory]['region2'],
            'region_label': factory_info[factory]['region_label'],
            'avg_rate':     avg_rate,
            'total_inspec': total_inspec,
            'total_defect': total_defect,
            'record_count': total_count,
            'trend':        trend,
            'monthly':      monthly_rates,
        })

    # 불량률 오름차순 정렬 (낮을수록 좋음), None은 끝으로
    result.sort(key=lambda x: (x['avg_rate'] is None, x['avg_rate'] or 9999))
    for i, r in enumerate(result):
        r['rank'] = i + 1

    return result


def _calc_trend(monthly_rates: list[dict]) -> str:
    """최근 3개월 이상 데이터가 있으면 기울기로 추이 판단"""
    valid = [m for m in monthly_rates[-3:] if m['rate'] is not None]
    if len(valid) < 2:
        return 'new'
    rates = [m['rate'] for m in valid]
    delta = rates[-1] - rates[0]
    if delta > 0.3:  return 'up'
    if delta < -0.3: return 'down'
    return 'flat'


# ── 지역별 히트맵 집계 ───────────────────────────────────────────
def calc_region_heatmap(raw_rows: list[dict],
                        start: str = None, end: str = None) -> list[dict]:
    """
    지역1 기준 집계
    반환: [{ region1, avg_rate, total_inspec, total_defect, factory_count }]
    """
    rows = _filter_rows(raw_rows, start, end)
    region: dict[str, dict] = defaultdict(lambda: {'inspec': 0, 'defect': 0, 'factories': set()})

    for r in rows:
        reg = r.get('region_label') or r.get('region1') or '(미입력)'
        region[reg]['inspec']    += r.get('inspec', 0) or 0
        region[reg]['defect']    += r.get('qty_total', 0) or 0
        f = r.get('factory', '')
        if f: region[reg]['factories'].add(f)

    result = []
    for reg, d in region.items():
        result.append({
            'region1':       reg,
            'avg_rate':      _safe_rate(d['defect'], d['inspec']),
            'total_inspec':  d['inspec'],
            'total_defect':  d['defect'],
            'factory_count': len(d['factories']),
        })
    result.sort(key=lambda x: (x['avg_rate'] is None, x['avg_rate'] or 9999))
    return result


# ── 공장 상세 드릴다운 ────────────────────────────────────────────
def calc_factory_detail(raw_rows: list[dict], cache: dict,
                        factory_name: str) -> dict:
    """특정 공장의 상세 분석 (월별 추이 + 불량 유형 TOP5)"""
    rows = [r for r in raw_rows if r.get('factory') == factory_name]
    if not rows:
        return {}

    # 월별 추이
    monthly: dict[str, dict] = defaultdict(lambda: {'inspec': 0, 'defect': 0})
    for r in rows:
        ym = _parse_date(r.get('date')) or 'unknown'
        monthly[ym]['inspec'] += r.get('inspec', 0) or 0
        monthly[ym]['defect'] += r.get('qty_total', 0) or 0

    monthly_list = [
        {'month': m, 'rate': _safe_rate(d['defect'], d['inspec']),
         'inspec': d['inspec'], 'defect': d['defect']}
        for m, d in sorted(monthly.items()) if m != 'unknown'
    ]

    # 불량 유형 TOP5 (표준불량명 기준)
    std_count: dict[str, int] = defaultdict(int)
    for r in rows:
        for (part, std, sc, meth, rev, note) in cache.get(r['defect_raw'], []):
            if std:
                std_count[std] += r.get('qty_total', 0) or 0
    top5 = sorted(std_count.items(), key=lambda x: -x[1])[:5]
    total_defect_sum = sum(std_count.values()) or 1

    # 바이어 목록
    buyers = list({r.get('buyer', '') for r in rows if r.get('buyer')})

    return {
        'factory':       factory_name,
        'region1':       rows[0].get('region1', ''),
        'region2':       rows[0].get('region2', ''),
        'buyers':        buyers,
        'record_count':  len(rows),
        'total_inspec':  sum(r.get('inspec', 0) or 0 for r in rows),
        'total_defect':  sum(r.get('qty_total', 0) or 0 for r in rows),
        'monthly':       monthly_list,
        'top5_defects': [
            {'name': name, 'qty': qty, 'pct': round(qty / total_defect_sum * 100, 1)}
            for name, qty in top5
        ],
    }


# ── AI 코멘트용 데이터 구성 ──────────────────────────────────────
def build_ai_comment_data(ranking: list[dict], period: str) -> dict:
    """Claude API에 넘길 요약 데이터 구성"""
    valid = [r for r in ranking if r['avg_rate'] is not None]
    best  = valid[:5]
    worst = valid[-5:][::-1]

    notable = [
        r for r in valid
        if r['trend'] == 'up' and r['avg_rate'] and r['avg_rate'] > 2.0
    ][:3]

    return {
        'period':       period,
        'total_factories': len(ranking),
        'top5_best': [
            {'factory': r['factory'], 'region': r.get('region_label') or r['region1'],
             'rate': r['avg_rate'], 'trend': r['trend']}
            for r in best
        ],
        'top5_worst': [
            {'factory': r['factory'], 'region': r.get('region_label') or r['region1'],
             'rate': r['avg_rate'], 'trend': r['trend']}
            for r in worst
        ],
        'notable_changes': [
            {'factory': r['factory'], 'region': r.get('region_label') or r['region1'],
             'rate': r['avg_rate'], 'direction': '급등'}
            for r in notable
        ],
    }


# ── 필터링 유틸 ─────────────────────────────────────────────────
def _filter_rows(rows, start=None, end=None, buyer=None, item=None):
    result = []
    for r in rows:
        ym = _parse_date(r.get('date'))
        if start and ym and ym < start: continue
        if end   and ym and ym > end:   continue
        if buyer and buyer != '전체' and r.get('buyer') != buyer: continue
        if item  and item  != '전체' and r.get('item')  != item:  continue
        result.append(r)
    return result


# ── 필터 옵션 추출 ───────────────────────────────────────────────
def get_filter_options(raw_rows: list[dict]) -> dict:
    buyers  = sorted({r.get('buyer','') for r in raw_rows if r.get('buyer')})
    items   = sorted({r.get('item','')  for r in raw_rows if r.get('item')})
    months  = sorted({_parse_date(r.get('date')) for r in raw_rows
                      if _parse_date(r.get('date'))})
    return {
        'buyers':  buyers,
        'items':   items,
        'months':  months,
        'start':   months[0]  if months else None,
        'end':     months[-1] if months else None,
    }
