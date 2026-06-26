"""
3_공장_지역_분석.py — 공장·지역별 불량률 분석
"""
import sys
from pathlib import Path

import streamlit as st
import pandas as pd
import plotly.express as px

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from core.factory_ranking import (
    calc_factory_ranking, calc_region_heatmap, calc_factory_detail,
    build_ai_comment_data, get_filter_options,
)
from core.ai_comment import get_comment
from core.pdf_report import generate_factory_pdf

st.set_page_config(page_title="공장·지역 분석", page_icon="🏭", layout="wide")
st.title("🏭 공장·지역 분석")

if "raw_rows" not in st.session_state or "cache" not in st.session_state:
    st.warning("먼저 '📊 불량명 표준화 매핑' 탭에서 데이터를 업로드하고 분석을 실행해주세요.")
    st.stop()

raw_rows = st.session_state.raw_rows
cache = st.session_state.cache

opts = get_filter_options(raw_rows)
months = opts['months']

st.markdown("### 필터")
c1, c2, c3, c4 = st.columns(4)
with c1:
    start = st.selectbox("시작 월", months, index=0 if months else None)
with c2:
    end = st.selectbox("종료 월", months, index=len(months) - 1 if months else None)
with c3:
    buyer = st.selectbox("바이어", ["전체"] + opts['buyers'])
with c4:
    item = st.selectbox("품명", ["전체"] + opts['items'])

ranking = calc_factory_ranking(raw_rows, cache, start=start, end=end, buyer=buyer, item=item)
heatmap = calc_region_heatmap(raw_rows, start=start, end=end)

trend_mark = {'up': '↑ 악화', 'down': '↓ 개선', 'flat': '→ 보합', 'new': '(데이터 부족)'}

st.markdown("### 🏆 공장별 불량률 랭킹")
if ranking:
    rank_df = pd.DataFrame(ranking)[
        ['rank', 'factory', 'region_label', 'avg_rate', 'total_inspec', 'total_defect', 'record_count', 'trend']
    ].rename(columns={
        'rank': '순위', 'factory': '공장', 'region_label': '지역',
        'avg_rate': '평균불량률(%)', 'total_inspec': '검사수량', 'total_defect': '불량수량',
        'record_count': '건수', 'trend': '추이',
    })
    rank_df['추이'] = rank_df['추이'].map(trend_mark)
    st.dataframe(rank_df, use_container_width=True, height=400)
else:
    st.info("선택한 조건에 해당하는 데이터가 없습니다.")

st.markdown("### 🗺️ 지역별 불량률 히트맵")
if heatmap:
    heat_df = pd.DataFrame(heatmap).rename(columns={
        'region1': '지역', 'avg_rate': '평균불량률(%)', 'total_inspec': '검사수량',
        'total_defect': '불량수량', 'factory_count': '공장수',
    })
    fig = px.bar(heat_df, x='지역', y='평균불량률(%)', color='평균불량률(%)',
                  color_continuous_scale='RdYlGn_r', text='평균불량률(%)')
    st.plotly_chart(fig, use_container_width=True)
    st.dataframe(heat_df, use_container_width=True)
else:
    st.info("지역별 데이터가 없습니다.")

st.markdown("### 💬 코멘트")
if ranking:
    period = f"{start} ~ {end}" if start and end else "전체기간"
    ai_data = build_ai_comment_data(ranking, period)
    try:
        api_key = st.secrets.get("ANTHROPIC_API_KEY")
    except Exception:
        api_key = None
    if st.button("코멘트 생성"):
        with st.spinner("코멘트 생성 중..."):
            comment = get_comment(ai_data, api_key)
        st.markdown(comment)

st.markdown("### 🔍 공장 상세")
if ranking:
    factory_names = [r['factory'] for r in ranking]
    selected_factory = st.selectbox("공장 선택", factory_names)

    if selected_factory:
        detail = calc_factory_detail(raw_rows, cache, selected_factory)
        if detail:
            d1, d2, d3 = st.columns(3)
            d1.metric("검사수량 합계", f"{detail['total_inspec']:,}")
            d2.metric("불량수량 합계", f"{detail['total_defect']:,}")
            d3.metric("데이터 건수", f"{detail['record_count']:,}")

            if detail['monthly']:
                m_df = pd.DataFrame(detail['monthly'])
                fig2 = px.line(m_df, x='month', y='rate', markers=True,
                                title="월별 불량률 추이(%)")
                st.plotly_chart(fig2, use_container_width=True)

            if detail['top5_defects']:
                st.markdown("**불량 유형 TOP5**")
                top5_df = pd.DataFrame(detail['top5_defects']).rename(
                    columns={'name': '불량명', 'qty': '수량', 'pct': '비율(%)'}
                )
                st.dataframe(top5_df, use_container_width=True)

            pdf_bytes = generate_factory_pdf(detail)
            st.download_button(
                "⬇️ 공장별 PDF 보고서 다운로드",
                data=pdf_bytes,
                file_name=f"{selected_factory}_분석보고서.pdf",
                mime="application/pdf",
            )
        else:
            st.info("해당 공장의 데이터가 없습니다.")
