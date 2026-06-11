"""
2_불량명_표준화.py — 불량명 표준화 매핑
"""
import sys, tempfile, shutil
from pathlib import Path

import streamlit as st
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from core.defect_core import (
    load_standard, load_raw, build_mapping, mapping_to_records,
    calc_stats, build_excel, save_corrections_to_std,
)

st.set_page_config(page_title="불량명 표준화 매핑", page_icon="📊", layout="wide")
st.title("📊 불량명 표준화 매핑")

DEFAULT_STD_PATH = Path(__file__).resolve().parent.parent / "표준불량명칭.xlsx"

if "tmpdir" not in st.session_state:
    st.session_state.tmpdir = tempfile.mkdtemp(prefix="defect_")
tmpdir = Path(st.session_state.tmpdir)

st.markdown("### 1단계: 표준불량명칭 파일")
std_file = st.file_uploader(
    "표준불량명칭.xlsx 업로드 (선택, 업로드하지 않으면 기본 파일 사용)", type=["xlsx"]
)

if std_file is not None:
    std_path = tmpdir / "표준불량명칭.xlsx"
    std_path.write_bytes(std_file.getvalue())
    st.session_state.std_path = str(std_path)
elif "std_path" not in st.session_state:
    std_path = tmpdir / "표준불량명칭.xlsx"
    shutil.copy(DEFAULT_STD_PATH, std_path)
    st.session_state.std_path = str(std_path)

std_path = Path(st.session_state.std_path)
st.caption(f"사용 중인 표준불량명칭 파일: {'업로드된 파일' if std_file is not None else '기본 파일'}")

st.markdown("### 2단계: 불량상세 데이터 업로드")
raw_files = st.file_uploader(
    "불량상세 데이터 (Excel, '② 불량상세' 시트 포함, 여러 개 선택 가능)",
    type=["xlsx"], accept_multiple_files=True,
)

if raw_files and st.button("🔍 매핑 분석 시작", type="primary"):
    with st.spinner("분석 중..."):
        raw_paths = []
        for f in raw_files:
            p = tmpdir / f.name
            p.write_bytes(f.getvalue())
            raw_paths.append(str(p))

        std_names, adict, used_sheet = load_standard(str(std_path))
        raw_rows, skipped = load_raw(raw_paths)
        cache, catmap = build_mapping(raw_rows, std_names, adict)

        st.session_state.std_names = std_names
        st.session_state.adict = adict
        st.session_state.raw_rows = raw_rows
        st.session_state.cache = cache
        st.session_state.catmap = catmap
        st.session_state.skipped = skipped

    st.success("분석 완료!")

if "raw_rows" in st.session_state:
    raw_rows = st.session_state.raw_rows
    cache = st.session_state.cache
    catmap = st.session_state.catmap
    std_names = st.session_state.std_names

    if st.session_state.get("skipped"):
        st.warning("건너뛴 파일: " + ", ".join(st.session_state.skipped))

    stats = calc_stats(cache)
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("전체 항목", stats['total'])
    c2.metric("자동매핑", stats['auto'])
    c3.metric("검토필요", stats['review'])
    c4.metric("미매핑", stats['unmapped'])
    c5.metric("자동매핑률", f"{stats['auto_pct']}%")

    records = mapping_to_records(raw_rows, cache, catmap)
    df = pd.DataFrame(records)

    st.markdown("### 매핑 결과")
    filter_opt = st.radio("필터", ["전체", "검토 필요", "미매핑"], horizontal=True)
    if filter_opt == "검토 필요":
        view = df[df['review'] == True]
    elif filter_opt == "미매핑":
        view = df[df['method'] == '미매핑']
    else:
        view = df

    show_cols = ['file', 'report_no', 'date', 'factory', 'defect_raw',
                  'part', 'std', 'category', 'score', 'method', 'review', 'note']
    st.dataframe(view[show_cols], use_container_width=True, height=400)

    st.markdown("### 수동 수정 — 검토 필요 / 미매핑 항목")
    review_df = (
        df[df['review'] == True][['part', 'std', 'method', 'score']]
        .drop_duplicates(subset=['part'])
        .reset_index(drop=True)
    )
    if not review_df.empty:
        review_df = review_df.rename(columns={
            'part': '분리불량명', 'std': '추천 표준명', 'method': '매핑방법', 'score': '신뢰도',
        })
        review_df['확정표준명'] = review_df['추천 표준명']
        std_options = [s[0] for s in std_names]

        edited = st.data_editor(
            review_df,
            column_config={
                "확정표준명": st.column_config.SelectboxColumn(options=[""] + std_options),
            },
            disabled=['분리불량명', '추천 표준명', '매핑방법', '신뢰도'],
            use_container_width=True,
            height=300,
            key="correction_editor",
        )

        if st.button("✅ 수정사항 적용 및 표준불량명칭에 저장"):
            corrections = []
            for _, row in edited.iterrows():
                if row['확정표준명'] and row['확정표준명'] != row['추천 표준명']:
                    corrections.append({"part": row['분리불량명'], "std": row['확정표준명']})
            if corrections:
                added = save_corrections_to_std(str(std_path), corrections)
                st.success(f"{added}개 별칭이 표준불량명칭.xlsx에 저장되었습니다. 재분석합니다...")
                std_names, adict, _ = load_standard(str(std_path))
                cache, catmap = build_mapping(raw_rows, std_names, adict)
                st.session_state.std_names = std_names
                st.session_state.adict = adict
                st.session_state.cache = cache
                st.session_state.catmap = catmap
                st.rerun()
            else:
                st.info("변경된 항목이 없습니다.")
    else:
        st.info("검토가 필요한 항목이 없습니다.")

    st.markdown("### 다운로드")
    col1, col2 = st.columns(2)
    with col1:
        out_path = tmpdir / "불량명_표준화_매핑결과.xlsx"
        build_excel(raw_rows, cache, std_names, catmap, str(out_path))
        st.download_button(
            "⬇️ 매핑 결과 Excel 다운로드",
            data=out_path.read_bytes(),
            file_name="불량명_표준화_매핑결과.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    with col2:
        st.download_button(
            "⬇️ 수정된 표준불량명칭.xlsx 다운로드",
            data=std_path.read_bytes(),
            file_name="표준불량명칭.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    st.info("💡 이 화면에서 분석한 데이터는 '🏭 공장·지역 분석' 탭에서 그대로 사용할 수 있습니다.")
elif not raw_files:
    st.info("불량상세 데이터 파일을 업로드 후 '매핑 분석 시작' 버튼을 눌러주세요.")
