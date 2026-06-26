"""
4_실적_분석.py — 제품평가팀 실적 분석 대시보드
"""
import sys, tempfile
from pathlib import Path

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from core.performance_core import (
    load_performance, filter_rows, get_filter_options,
    summary_by_year, summary_by_region, summary_by_buyer, summary_by_brand,
    summary_by_code, region_code_crosstab, yoy_comparison, monthly_compare,
    cumulative_by_year, actual_by_month_code, load_plan_budget, CODE_LABELS,
)

st.set_page_config(page_title="실적 분석", page_icon="📈", layout="wide")
st.title("📈 실적 분석")

DEFAULT_PLAN_PATH = Path(__file__).resolve().parent.parent / "plan_budget.xlsx"


def fmt_money(v):
    return f"{v/1e8:,.2f}억원" if abs(v) >= 1e8 else f"{v:,.0f}원"


if "tmpdir" not in st.session_state:
    st.session_state.tmpdir = tempfile.mkdtemp(prefix="defect_")
tmpdir = Path(st.session_state.tmpdir)

st.markdown("### 데이터 업로드")
perf_file = st.file_uploader("실적 rawdata 업로드 (Excel)", type=["xlsx"])

if perf_file is not None and st.button("📥 데이터 로드", type="primary"):
    with st.spinner("로드 중..."):
        p = tmpdir / perf_file.name
        p.write_bytes(perf_file.getvalue())
        rows = load_performance(str(p))
        st.session_state.perf_rows = rows
    st.success(f"{len(rows):,}건 로드 완료")

if "perf_rows" not in st.session_state:
    st.info("실적 rawdata 파일을 업로드하고 '데이터 로드' 버튼을 눌러주세요.")
    st.stop()

rows = st.session_state.perf_rows
opts = get_filter_options(rows)

st.markdown("### 필터")
c1, c2, c3, c4 = st.columns(4)
with c1:
    f_years = st.multiselect("연도", opts['years'], default=opts['years'])
with c2:
    f_region = st.selectbox("지역", ["전체"] + opts['regions'])
with c3:
    f_code = st.selectbox("코드", ["전체"] + [f"{c} ({CODE_LABELS.get(c, c)})" for c in opts['codes']])
    f_code_val = f_code.split(" ")[0] if f_code != "전체" else None
with c4:
    f_buyer = st.selectbox("바이어", ["전체"] + opts['buyers'])

frows = filter_rows(
    rows,
    years=f_years if f_years else None,
    region=None if f_region == "전체" else f_region,
    code=f_code_val,
    buyer=None if f_buyer == "전체" else f_buyer,
)

# ── KPI ──────────────────────────────────────────────────────────
st.markdown("### KPI")
yearly = summary_by_year(frows)
if yearly:
    cols = st.columns(len(yearly))
    for col, y in zip(cols, yearly):
        col.metric(f"{y['year']}년 수익", fmt_money(y['revenue']), f"{y['cnt']:,}건")
else:
    st.info("선택한 조건에 해당하는 데이터가 없습니다.")

# ── 3개년 월별 추이 ──────────────────────────────────────────────
st.markdown("### 📊 3개년 월별 추이")
dim_label = st.selectbox("비교 기준", ["전체", "지역별", "바이어별", "브랜드별", "코드별"])

dim = None
group_key = None
if dim_label == "지역별":
    dim = "region"
    group_key = st.selectbox("지역 선택", opts['regions'])
elif dim_label == "바이어별":
    dim = "buyer"
    group_key = st.selectbox("바이어 선택", opts['buyers'])
elif dim_label == "브랜드별":
    dim = "brand"
    brand_opts = [b['brand'] for b in summary_by_brand(frows, top_n=30)]
    group_key = st.selectbox("브랜드 선택", brand_opts)
elif dim_label == "코드별":
    dim = "code"
    code_opts = {f"{c} ({CODE_LABELS.get(c, c)})": CODE_LABELS.get(c, c) for c in opts['codes']}
    sel = st.selectbox("코드 선택", list(code_opts.keys()))
    group_key = code_opts[sel]

years_for_trend = tuple(sorted(f_years)[-3:]) if f_years else (2024, 2025, 2026)
mc = monthly_compare(frows, dim=dim, group_filter={group_key} if group_key else None, years=years_for_trend)

target_group = group_key if group_key else "전체"
data = mc.get(target_group, {})
if data:
    plot_rows = []
    for y, mms in data.items():
        for mm, rev in mms.items():
            plot_rows.append({'연도': str(y), '월': mm, '수익': rev})
    pdf = pd.DataFrame(plot_rows).sort_values(['연도', '월'])
    fig = px.line(pdf, x='월', y='수익', color='연도', markers=True,
                   title=f"{target_group} — 월별 수익 추이")
    st.plotly_chart(fig, use_container_width=True)
else:
    st.info("표시할 데이터가 없습니다.")

# ── 동기누적 비교 ─────────────────────────────────────────────────
st.markdown("### 🔁 동기누적 비교")
latest_year = max(opts['years']) if opts['years'] else None
if latest_year:
    latest_months = sorted({
        r['ym'][-2:] for r in rows if r['year'] == latest_year and r.get('ym') and len(r['ym']) == 7
    })
    if latest_months:
        st.caption(f"기준: {latest_year}년 {latest_months[0]}~{latest_months[-1]}월 (동기간 누적)")
        cum_years = tuple(sorted({y for y in (latest_year - 2, latest_year - 1, latest_year)}))
        cum = cumulative_by_year(frows, latest_months, years=cum_years)
        cum_df = pd.DataFrame([
            {'연도': y, '누적건수': v['cnt'], '누적수익': v['rev'], '표시': fmt_money(v['rev'])}
            for y, v in cum.items()
        ])
        st.dataframe(cum_df[['연도', '누적건수', '표시']].rename(columns={'표시': '누적수익'}),
                      use_container_width=True)

        yoy = yoy_comparison(frows, dim=dim, same_months=latest_months, top_n=10, sort_year=latest_year)
        if yoy:
            yoy_df = pd.DataFrame(yoy)
            display_cols = ['label'] + [c for c in yoy_df.columns if c.startswith('y')] + ['growth_24_25', 'growth_25_26']
            display_cols = [c for c in display_cols if c in yoy_df.columns]
            st.dataframe(yoy_df[display_cols], use_container_width=True)
    else:
        st.info("최신 연도의 월별 데이터가 없습니다.")

# ── 목표 vs 실적 (131/152) ───────────────────────────────────────
st.markdown("### 🎯 목표 대비 실적 (131/152)")
plan_file = st.file_uploader("목표예산 파일 업로드 (선택, 미업로드시 기본 파일 사용)", type=["xlsx"], key="plan_upload")
if plan_file is not None:
    plan_path = tmpdir / "plan_budget.xlsx"
    plan_path.write_bytes(plan_file.getvalue())
else:
    plan_path = DEFAULT_PLAN_PATH

if latest_year:
    try:
        plan = load_plan_budget(str(plan_path))
        actual = actual_by_month_code(rows, latest_year)
        plan_rows = []
        for mm in [f"{i:02d}" for i in range(1, 13)]:
            p = plan['monthly'].get(mm, {'131': 0, '152': 0, 'total': 0})
            a = actual.get(mm, {'131': 0, '152': 0, 'total': 0})
            plan_rows.append({
                '월': mm,
                '목표(131)': p['131'], '실적(131)': a['131'],
                '목표(152)': p['152'], '실적(152)': a['152'],
                '목표(합계)': p['total'], '실적(합계)': a['total'],
                '달성률(%)': round(a['total'] / p['total'] * 100, 1) if p['total'] else None,
            })
        plan_df = pd.DataFrame(plan_rows)

        fig2 = go.Figure()
        fig2.add_bar(x=plan_df['월'], y=plan_df['목표(합계)'], name='목표')
        fig2.add_bar(x=plan_df['월'], y=plan_df['실적(합계)'], name='실적')
        fig2.update_layout(barmode='group', title=f"{latest_year}년 월별 목표 대비 실적 (합계)")
        st.plotly_chart(fig2, use_container_width=True)

        st.dataframe(plan_df, use_container_width=True)
    except Exception as e:
        st.warning(f"목표예산 파일을 처리할 수 없습니다: {e}")

# ── 지역 × 코드 교차분석 ──────────────────────────────────────────
st.markdown("### 🗺️ 지역 × 코드 교차분석")
crosstab = region_code_crosstab(frows)
if crosstab:
    ct_df = pd.DataFrame(crosstab).rename(columns={
        'region': '지역', 'c131': '131(원단검사)', 'c152': '152(완제품검사)',
        'other': '기타', 'total': '합계',
    })
    fig3 = go.Figure()
    fig3.add_bar(x=ct_df['지역'], y=ct_df['131(원단검사)'], name='131(원단검사)')
    fig3.add_bar(x=ct_df['지역'], y=ct_df['152(완제품검사)'], name='152(완제품검사)')
    fig3.update_layout(barmode='stack', title="지역별 코드 구성")
    st.plotly_chart(fig3, use_container_width=True)
    st.dataframe(ct_df, use_container_width=True)
else:
    st.info("표시할 데이터가 없습니다.")
