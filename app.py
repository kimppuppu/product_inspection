"""
app.py — 제품평가팀 불량률·실적 분석 웹앱 (Streamlit, FITI UI)
실행: streamlit run app.py
"""
import sys
import shutil
import tempfile
from pathlib import Path
from datetime import datetime
from collections import defaultdict

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

sys.path.insert(0, str(Path(__file__).resolve().parent))

_IMPORT_ERROR = None
try:
    from core.pdf_extractor import parse_pdf, make_workbook
    from core.defect_core import (
        load_standard, load_raw, build_mapping, mapping_to_records,
        calc_stats, build_excel, save_corrections_to_std,
        load_standard_typed, build_mapping_typed, classify_item_type,
    )
    from core.factory_ranking import (
        calc_factory_ranking, calc_region_heatmap, calc_factory_detail,
        build_ai_comment_data, get_filter_options as get_factory_filter_options,
    )
    from core.ai_comment import get_comment
    from core.pdf_report import generate_factory_pdf
    from core.word_report import generate_word_report
    from core.performance_core import (
        load_performance, filter_rows, get_filter_options as get_perf_filter_options,
        summary_by_brand, region_code_crosstab, yoy_comparison, monthly_compare,
        cumulative_by_year, actual_by_month_code, load_plan_budget, CODE_LABELS,
        summary_by_year, REGION_ORDER,
    )
    _IMPORT_ERROR = None
except Exception as _e:
    import traceback
    _IMPORT_ERROR = traceback.format_exc()

st.set_page_config(
    page_title="제품평가 업무관리",
    page_icon="📋",
    layout="wide",
)

if _IMPORT_ERROR:
    st.error("앱 시작 오류 — 아래 내용을 개발자에게 전달해주세요:")
    st.code(_IMPORT_ERROR)
    st.stop()

# ──────────────────────────────────────────────────────────────────
# FITI 브랜딩 — 커스텀 CSS
# ──────────────────────────────────────────────────────────────────
st.markdown("""
<style>
:root {
    --primary: #0052a3;
    --primary-light: #e8f3fc;
    --fiti-blue: #0075c9;
    --success: #1e7e34;
    --success-bg: #C6EFCE;
    --warn-bg: #FFEB9C;
    --danger-bg: #FFC7CE;
    --gray: #6c757d;
    --border: #dee2e6;
    --bg: #f2f6fb;
}

/* 전체 배경 */
.stApp { background: var(--bg); }

/* 상단 여백 축소 + 폭 */
.block-container { padding-top: 0.5rem; padding-bottom: 2rem; max-width: 1400px; }

/* 사이드바 숨김 (단일 화면 + 상단 탭 구조) */
[data-testid="stSidebar"], [data-testid="collapsedControl"] { display: none; }

/* Streamlit 기본 상단 헤더/툴바 숨김 (FITI 헤더가 가려지는 문제 해결) */
[data-testid="stHeader"], [data-testid="stToolbar"] { display: none !important; height: 0 !important; }

/* FITI 헤더 바 */
.fiti-header {
    background: #003f85;
    color: white;
    padding: 14px 28px;
    margin: 0 -1rem 1.2rem -1rem;
    border-bottom: 3px solid var(--fiti-blue);
    display: flex;
    align-items: center;
    box-shadow: 0 2px 12px rgba(0,0,0,0.18);
}
.fiti-logo-wrap { display: flex; align-items: center; gap: 14px; }
.fiti-logo-block { display: flex; align-items: center; gap: 10px; }
.fiti-logo-text { font-size: 26px; font-weight: 900; letter-spacing: -1px; line-height: 1; }
.fiti-logo-text .fi { color: #ffffff; }
.fiti-logo-text .ti { color: var(--fiti-blue); }
.fiti-logo-sub { display: flex; flex-direction: column; line-height: 1.3; }
.fiti-logo-sub-kr { font-size: 13px; font-weight: 700; color: #ffffff; }
.fiti-logo-sub-en { font-size: 9px; font-weight: 400; color: rgba(255,255,255,0.6); letter-spacing: 0.2px; }
.fiti-divider { width: 1px; height: 34px; background: rgba(255,255,255,0.25); }
.fiti-app-name { font-size: 17px; font-weight: 800; letter-spacing: -0.3px; }
.fiti-app-sub { font-size: 11px; opacity: 0.65; margin-top: 2px; letter-spacing: 0.2px; }

/* 상단 탭 — index.html .tabs / .tab 스타일 매칭 */
.stTabs [data-baseweb="tab-list"] {
    gap: 4px;
    background: white;
    border-bottom: 2px solid var(--border);
    padding: 0 12px;
    border-radius: 10px 10px 0 0;
    box-shadow: 0 2px 8px rgba(0,0,0,0.06);
}
.stTabs [data-baseweb="tab"] {
    height: 48px;
    font-weight: 600;
    font-size: 14px;
    color: var(--gray);
}
.stTabs [data-baseweb="tab"]:hover { color: var(--primary); }
.stTabs [aria-selected="true"] {
    color: var(--primary) !important;
    border-bottom: 3px solid var(--fiti-blue) !important;
}
.stTabs [data-baseweb="tab-panel"] { padding-top: 1.2rem; }

/* 패널/섹션 제목 — index.html .panel-title 매칭 */
.panel-title {
    font-size: 15px; font-weight: 700; color: var(--primary);
    margin: 1.4rem 0 0.9rem 0; padding-bottom: 8px;
    border-bottom: 1px solid var(--border);
}
.panel-title:first-child { margin-top: 0; }

/* 카드(KPI/지표) — index.html .card 스타일 매칭 */
[data-testid="stMetric"] {
    background: white;
    border-radius: 10px;
    padding: 16px 18px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.08);
}
[data-testid="stMetricValue"] { color: var(--primary); font-weight: 800; }
[data-testid="stMetricLabel"] { color: var(--gray); }

/* 버튼 — primary 색상은 config.toml과 함께 적용 */
.stButton>button[kind="primary"], .stDownloadButton>button[kind="primary"] {
    background: var(--primary);
    border-color: var(--primary);
}
.stButton>button[kind="primary"]:hover, .stDownloadButton>button[kind="primary"]:hover {
    background: #003f85;
    border-color: #003f85;
}

/* 안내 박스 / 데이터프레임 라운딩 */
[data-testid="stDataFrame"] { border-radius: 8px; overflow: hidden; }
</style>
""", unsafe_allow_html=True)

# ──────────────────────────────────────────────────────────────────
# FITI 헤더
# ──────────────────────────────────────────────────────────────────
st.markdown("""
<div class="fiti-header">
  <div class="fiti-logo-wrap">
    <div class="fiti-logo-block">
      <span class="fiti-logo-text"><span class="fi">FI</span><span class="ti">TI</span></span>
      <div class="fiti-logo-sub">
        <span class="fiti-logo-sub-kr">FITI 시험연구원</span>
        <span class="fiti-logo-sub-en">FITI Testing &amp; Research Institute</span>
      </div>
    </div>
    <div class="fiti-divider"></div>
    <div>
      <div class="fiti-app-name">제품평가 업무관리</div>
      <div class="fiti-app-sub">Product Quality Evaluation Management System</div>
    </div>
  </div>
</div>
""", unsafe_allow_html=True)


def panel_title(text: str):
    st.markdown(f'<div class="panel-title">{text}</div>', unsafe_allow_html=True)


def fmt_money(v):
    return f"{v/1e8:,.2f}억원" if abs(v) >= 1e8 else f"{v:,.0f}원"


def fmt_won(v):
    """천단위 콤마 + '원' 단위 표시 (표 표시용, 억원 약식 변환 없이 전체 자리수 표시)"""
    try:
        return f"{int(round(v)):,}원"
    except (TypeError, ValueError):
        return v


def fmt_won_kr(v):
    """억/만원 단위의 한국어 표기로 변환 (차트 축/툴팁 라벨용, 예: 7억원, 7,500만원, 320만원)"""
    try:
        v = int(round(v))
    except (TypeError, ValueError):
        return str(v)
    sign = '-' if v < 0 else ''
    v = abs(v)
    if v == 0:
        return "0원"
    eok, rem = divmod(v, 100_000_000)
    man = rem // 10_000
    parts = []
    if eok:
        parts.append(f"{eok}억")
    if man:
        parts.append(f"{man:,}만")
    if not parts:
        return f"{sign}{v:,}원"
    return f"{sign}{''.join(parts)}원"


def fmt_pct1(v):
    """소수점 첫째자리까지 % 표시"""
    try:
        return f"{float(v):.1f}%"
    except (TypeError, ValueError):
        return "-"


DEFAULT_STD_PATH = Path(__file__).resolve().parent / "표준불량명칭.xlsx"
DATA_DIR = Path(__file__).parent / "data"
DEFAULT_PLAN_PATH = Path(__file__).resolve().parent / "plan_budget.xlsx"

if "tmpdir" not in st.session_state or not Path(st.session_state.tmpdir).exists():
    st.session_state.tmpdir = tempfile.mkdtemp(prefix="defect_")
tmpdir = Path(st.session_state.tmpdir)


def run_mapping_analysis(raw_paths):
    """표준불량명칭/불량상세 데이터를 읽어 매핑을 수행하고 세션에 저장합니다."""
    raw_rows, skipped = load_raw(raw_paths)

    # 의류/잡화/신발 분류별 표준불량명칭 로드
    std_by_type = load_standard_typed(str(DATA_DIR))

    # 분류별 typed 매핑 (raw_rows에 product_type 필드 추가됨)
    cache, catmap = build_mapping_typed(raw_rows, std_by_type)

    # 수동수정 UI용: 전체 std_names 합본 (의류 기준, 나머지 추가)
    all_names, all_adict = [], {}
    for ptype in ['의류', '잡화', '신발']:
        if ptype in std_by_type:
            for item in std_by_type[ptype][0]:
                if item not in all_names:
                    all_names.append(item)
            all_adict.update(std_by_type[ptype][1])

    st.session_state.std_by_type = std_by_type
    st.session_state.std_names = all_names
    st.session_state.adict = all_adict
    st.session_state.raw_rows = raw_rows
    st.session_state.cache = cache
    st.session_state.catmap = catmap
    st.session_state.skipped = skipped


# ──────────────────────────────────────────────────────────────────
# 탭1: PDF → Excel 변환
# ──────────────────────────────────────────────────────────────────
def render_pdf_tab():
    import gc

    panel_title("📄 불량보고서 PDF → Excel 변환")

    BATCH = 50  # 한 번에 올릴 최대 개수

    if "pdf_records" not in st.session_state:
        st.session_state.pdf_records = []
    if "pdf_failed" not in st.session_state:
        st.session_state.pdf_failed = []
    if "pdf_seen" not in st.session_state:
        st.session_state.pdf_seen = set()
    if "pdf_upload_key" not in st.session_state:
        st.session_state.pdf_upload_key = 0

    total_rec = len(st.session_state.pdf_records)
    total_fail = len(st.session_state.pdf_failed)

    # 누적 현황 항상 상단 표시
    if total_rec + total_fail > 0:
        st.success(f"✅ 누적 {total_rec}개 변환 완료" + (f"  |  ❌ 실패 {total_fail}개" if total_fail else ""))

    st.info(f"📌 **{BATCH}개씩** 나눠서 올려주세요. 올릴 때마다 자동으로 처리 후 업로더가 초기화됩니다.")

    new_pdfs = st.file_uploader(
        f"PDF 파일 선택 (최대 {BATCH}개)",
        type=["pdf"],
        accept_multiple_files=True,
        key=f"pdf_uploader_{st.session_state.pdf_upload_key}",
    )

    if new_pdfs:
        new_files = [f for f in new_pdfs if f.name not in st.session_state.pdf_seen]
        if len(new_files) > BATCH:
            st.warning(f"⚠️ {len(new_files)}개 선택됐는데 앞 {BATCH}개만 처리합니다. 나머지는 다음 배치에 올려주세요.")
            new_files = new_files[:BATCH]

        if new_files:
            prog = st.progress(0.0)
            status = st.empty()
            with tempfile.TemporaryDirectory() as tdir:
                for i, f in enumerate(new_files, 1):
                    if total_rec + total_fail >= 1000:
                        st.warning("최대 1000개 한도에 도달했습니다.")
                        break
                    status.write(f"({i}/{len(new_files)}) 처리 중: {f.name}")
                    tmp_path = Path(tdir) / f.name
                    tmp_path.write_bytes(f.getvalue())
                    try:
                        rec = parse_pdf(str(tmp_path))
                        st.session_state.pdf_records.append(rec)
                    except Exception:
                        st.session_state.pdf_failed.append(f.name)
                    st.session_state.pdf_seen.add(f.name)
                    tmp_path.unlink(missing_ok=True)
                    gc.collect()
                    prog.progress(i / len(new_files))

            # 업로더 리셋 → 이전 배치 bytes 메모리 해제
            st.session_state.pdf_upload_key += 1
            st.rerun()

    # 하단 버튼 영역
    if total_rec + total_fail > 0:
        col_a, col_b = st.columns([3, 1])
        with col_a:
            make_btn = st.button("📥 Excel 파일 생성", type="primary", key="pdf_make_btn",
                                 disabled=(total_rec == 0))
        with col_b:
            if st.button("🗑️ 초기화", key="pdf_clear_btn"):
                st.session_state.pdf_records = []
                st.session_state.pdf_failed = []
                st.session_state.pdf_seen = set()
                st.session_state.pdf_convert_result = None
                st.session_state.pdf_upload_key += 1
                st.rerun()
        if total_fail > 0:
            with st.expander(f"실패 목록 ({total_fail}개)"):
                for name in st.session_state.pdf_failed:
                    st.caption(f"❌ {name}")
    else:
        make_btn = False

    if make_btn and total_rec > 0:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_name = f"불량율_분석_통합_{ts}.xlsx"
        out_path = tmpdir / out_name
        make_workbook(st.session_state.pdf_records, str(out_path))
        st.session_state.pdf_convert_result = {
            "out_path": str(out_path),
            "out_name": out_name,
            "success_count": total_rec,
            "failed": st.session_state.pdf_failed,
        }
        st.session_state.pdf_analysis_done = False

    result = st.session_state.get("pdf_convert_result")
    if result:
        panel_title("변환 결과")
        st.success(f"✅ 완료 — 성공 {result['success_count']}개 / 실패 {len(result['failed'])}개")

        out_path = Path(result["out_path"])
        if out_path.exists():
            col1, col2 = st.columns(2)
            with col1:
                with open(result["out_path"], "rb") as fp:
                    st.download_button(
                        "⬇️ Excel 다운로드",
                        data=fp,
                        file_name=result["out_name"],
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dl_excel",
                    )
            with col2:
                if st.button("📊 이 파일로 불량명 분석 시작", key="auto_analyze_btn"):
                    with st.spinner("분석 중..."):
                        run_mapping_analysis([result["out_path"]])
                    st.session_state.pdf_analysis_done = True

        if st.session_state.get("pdf_analysis_done"):
            st.success("✅ 분석이 완료되었습니다. '📊 불량명 표준화' 탭에서 결과를 확인하세요.")


def render_defect_tab():
    panel_title("📊 불량명 표준화 매핑")

    panel_title("1단계: 표준불량명칭 파일")
    std_file = st.file_uploader(
        "표준불량명칭.xlsx 업로드 (선택, 업로드하지 않으면 기본 파일 사용)",
        type=["xlsx"], key="std_uploader",
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

    panel_title("2단계: 불량상세 데이터 업로드")

    MAX_FILES = 1000

    # 누적 파일 저장소 초기화
    if "accumulated_files" not in st.session_state:
        st.session_state.accumulated_files = {}  # {filename: bytes}

    new_files = st.file_uploader(
        "불량상세 데이터 (Excel, '② 불량상세' 시트 포함, 여러 번 나눠서 추가 가능)",
        type=["xlsx"], accept_multiple_files=True, key="raw_uploader",
    )

    # 새로 업로드된 파일을 누적 목록에 추가
    if new_files:
        added = 0
        skipped_dup = []
        for f in new_files:
            if len(st.session_state.accumulated_files) >= MAX_FILES:
                st.warning(f"최대 {MAX_FILES}개 파일 한도에 도달했습니다.")
                break
            if f.name not in st.session_state.accumulated_files:
                st.session_state.accumulated_files[f.name] = f.getvalue()
                added += 1
            else:
                skipped_dup.append(f.name)
        if added:
            st.toast(f"{added}개 파일 추가됨 (누적: {len(st.session_state.accumulated_files)}개)")
        if skipped_dup:
            st.caption(f"중복 파일 무시: {', '.join(skipped_dup)}")

    total_acc = len(st.session_state.accumulated_files)
    if total_acc > 0:
        st.info(f"📂 누적 파일: **{total_acc}개** (최대 {MAX_FILES}개) — 파일을 추가한 후 분석을 실행하세요.")
        with st.expander(f"누적 파일 목록 ({total_acc}개)"):
            for name in sorted(st.session_state.accumulated_files.keys()):
                st.caption(f"• {name}")
        col_a, col_b = st.columns([3, 1])
        with col_a:
            run_btn = st.button("🔍 매핑 분석 시작", type="primary", key="mapping_btn",
                                disabled=(total_acc == 0))
        with col_b:
            if st.button("🗑️ 목록 초기화", key="clear_files_btn"):
                st.session_state.accumulated_files = {}
                st.rerun()
    else:
        st.info("파일을 업로드해주세요. 여러 번 나눠서 추가할 수 있습니다.")
        run_btn = False

    if run_btn and total_acc > 0:
        with st.spinner(f"{total_acc}개 파일 분석 중..."):
            raw_paths = []
            for fname, fbytes in st.session_state.accumulated_files.items():
                p = tmpdir / fname
                p.write_bytes(fbytes)
                raw_paths.append(str(p))

            run_mapping_analysis(raw_paths)

        st.success(f"분석 완료! ({total_acc}개 파일 처리)")

    if "raw_rows" in st.session_state:
        raw_rows = st.session_state.raw_rows
        cache = st.session_state.cache
        catmap = st.session_state.catmap
        std_names = st.session_state.std_names

        if st.session_state.get("skipped"):
            st.warning("건너뛴 파일: " + ", ".join(st.session_state.skipped))

        records = mapping_to_records(raw_rows, cache, catmap)
        df = pd.DataFrame(records)

        # 통계: 실제 행 기준, 미매핑 / 검토필요 / 자동매핑 명확히 분리
        _total  = len(df)
        _unmap  = int((df['method'] == '미매핑').sum())
        _review = int(((df['review'] == True) & (df['method'] != '미매핑')).sum())
        _auto   = _total - _unmap - _review
        _pct    = round(_auto / _total * 100, 1) if _total else 0
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("전체 항목", _total)
        c2.metric("자동매핑", _auto)
        c3.metric("검토필요", _review)
        c4.metric("미매핑", _unmap)
        c5.metric("자동매핑률", f"{_pct}%")

        panel_title("매핑 결과")
        filter_opt = st.radio("필터", ["전체", "검토 필요", "미매핑"], horizontal=True, key="mapping_filter")
        if filter_opt == "검토 필요":
            # 미매핑 제외, 퍼지매핑 저신뢰도만
            view = df[(df['review'] == True) & (df['method'] != '미매핑')]
        elif filter_opt == "미매핑":
            view = df[df['method'] == '미매핑']
        else:
            view = df

        # 컬럼 순서: defect_raw >> part >> 구분 >> 카테고리 >> 표준불량명
        display_cols = ['file', 'report_no', 'date', 'factory',
                        'defect_raw', 'part', 'product_type', 'category', 'std',
                        'score', 'method']

        # 구분(의류/잡화/신발)별 카테고리·표준불량명 옵션 분리 구성
        std_by_type_sess = st.session_state.get('std_by_type', {})
        _type_opts = {}  # ptype → (categories, cat_map, display_opts, name_to_display)
        for _pt in ['의류', '잡화', '신발']:
            _src = std_by_type_sess.get(_pt) or std_by_type_sess.get('의류')
            if not _src:
                continue
            _pnames, _ = _src
            _cats = list(dict.fromkeys([sc for _, sc, _ in _pnames if sc]))
            _cmap, _sbycat = {}, {}
            for sname, scat, _ in _pnames:
                lbl = f"[{scat}] {sname}" if scat else sname
                _cmap[lbl] = sname
                _sbycat.setdefault(scat or '기타', []).append(lbl)
            _dopts = []
            for cat in _cats:
                _dopts.extend(_sbycat.get(cat, []))
            _type_opts[_pt] = (_cats, _cmap, _dopts, {v: k for k, v in _cmap.items()})

        # view에 원본값 추적 컬럼 추가 (editor에서는 column_order로 숨김)
        show_cols_all = ['file', 'report_no', 'date', 'factory',
                         'defect_raw', 'part', 'product_type', 'category', 'std',
                         'score', 'method', 'review', 'note']
        view_edit = view[show_cols_all].reset_index(drop=True).copy()
        view_edit['_orig_std']  = view['std'].reset_index(drop=True).values
        view_edit['_orig_part'] = view['part'].reset_index(drop=True).values

        st.caption("💡 구분 탭에서 의류/잡화/신발을 선택 → 카테고리 드롭다운 → 표준불량명 드롭다운 순으로 선택하세요.")

        ptypes_in_view = [pt for pt in ['의류', '잡화', '신발']
                          if pt in view_edit['product_type'].values]
        all_edited = {}  # ptype → (edited_df, cat_map)

        if not ptypes_in_view:
            st.info("표시할 항목이 없습니다.")
        elif len(ptypes_in_view) == 1:
            _containers = [st.container()]
        else:
            _containers = st.tabs(ptypes_in_view)

        for _tc, _pt in zip(_containers, ptypes_in_view):
            with _tc:
                if _pt not in _type_opts:
                    st.warning(f"{_pt} 표준불량명칭 파일이 없습니다.")
                    continue
                _cats, _cmap, _dopts, _n2d = _type_opts[_pt]
                _pdf = view_edit[view_edit['product_type'] == _pt].reset_index(drop=True).copy()
                _pdf['std'] = _pdf['_orig_std'].apply(lambda n: _n2d.get(n, n) if n else "")
                if _pdf.empty:
                    st.info(f"{_pt} 항목 없음")
                    continue
                _edited = st.data_editor(
                    _pdf,
                    column_order=display_cols,
                    column_config={
                        "product_type": st.column_config.TextColumn(label="구분"),
                        "category": st.column_config.SelectboxColumn(
                            label="카테고리",
                            options=[""] + _cats,
                            help=f"{_pt} 카테고리를 선택하세요",
                        ),
                        "std": st.column_config.SelectboxColumn(
                            label="표준불량명",
                            options=[""] + _dopts,
                            help="카테고리명 입력 시 빠른 검색 가능 (예: '봉제')",
                        ),
                    },
                    disabled=[c for c in display_cols if c not in ('category', 'std')],
                    use_container_width=True,
                    height=400,
                    key=f"defect_editor_{_pt}",
                )
                all_edited[_pt] = (_edited, _cmap)

        if st.button("✅ 표준불량명 저장 및 재분석", key="save_correction_btn"):
            corrections = []
            for _pt, (_edited, _cmap) in all_edited.items():
                for _, row in _edited.iterrows():
                    sel_disp = row['std']
                    sel_name = _cmap.get(sel_disp, sel_disp) if sel_disp else ""
                    orig_name = row.get('_orig_std', "")
                    part_name = row.get('_orig_part', "")
                    if sel_name and sel_name != orig_name:
                        corrections.append({"part": part_name, "std": sel_name})
            if corrections:
                added = save_corrections_to_std(str(std_path), corrections)
                st.success(f"{added}개 별칭이 표준불량명칭.xlsx에 저장되었습니다. 재분석합니다...")
                std_by_type = load_standard_typed(str(DATA_DIR))
                cache, catmap = build_mapping_typed(raw_rows, std_by_type)
                all_names, all_adict = [], {}
                for ptype in ['의류', '잡화', '신발']:
                    if ptype in std_by_type:
                        for item in std_by_type[ptype][0]:
                            if item not in all_names:
                                all_names.append(item)
                        all_adict.update(std_by_type[ptype][1])
                st.session_state.std_by_type = std_by_type
                st.session_state.std_names = all_names
                st.session_state.adict = all_adict
                st.session_state.cache = cache
                st.session_state.catmap = catmap
                st.rerun()
            else:
                st.info("변경된 항목이 없습니다.")

        panel_title("다운로드")
        col1, col2 = st.columns(2)
        with col1:
            out_path = tmpdir / "불량명_표준화_매핑결과.xlsx"
            build_excel(raw_rows, cache, std_names, catmap, str(out_path))
            st.download_button(
                "⬇️ 매핑 결과 Excel 다운로드",
                data=out_path.read_bytes(),
                file_name="불량명_표준화_매핑결과.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_mapping",
            )
        with col2:
            st.download_button(
                "⬇️ 수정된 표준불량명칭.xlsx 다운로드",
                data=std_path.read_bytes(),
                file_name="표준불량명칭.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_std",
            )

        st.markdown("---")
        panel_title("📄 불량률 분석 보고서 (Word)")
        st.markdown("전체·업체별·공장별 불량률과 세부 불량 유형을 포함한 Word 보고서를 생성합니다.  \n※ 상위 5개 불량 유형 + 기타 · 차트 포함")
        try:
            word_bytes = generate_word_report(raw_rows, cache)
            today = datetime.now().strftime("%Y%m%d")
            st.download_button(
                "📄 Word 보고서 다운로드",
                data=word_bytes,
                file_name=f"불량률분석보고서_{today}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key="dl_word",
            )
        except Exception as e:
            st.error(f"보고서 생성 오류: {e}")

        st.info("💡 이 화면에서 분석한 데이터는 '🏭 공장·지역 분석' 탭에서 그대로 사용할 수 있습니다.")
    elif total_acc == 0:
        st.info("불량상세 데이터 파일을 업로드 후 '매핑 분석 시작' 버튼을 눌러주세요.")


# ──────────────────────────────────────────────────────────────────
# 탭3: 공장·지역 분석
# ──────────────────────────────────────────────────────────────────
def render_factory_tab():
    panel_title("🏭 공장·지역 분석")

    if "raw_rows" not in st.session_state or "cache" not in st.session_state:
        st.warning("먼저 '📊 불량명 표준화' 탭에서 데이터를 업로드하고 분석을 실행해주세요.")
        return

    raw_rows = st.session_state.raw_rows
    cache = st.session_state.cache

    opts = get_factory_filter_options(raw_rows)
    months = opts['months']

    panel_title("필터")
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        start = st.selectbox("시작 월", months, index=0 if months else None, key="f_start")
    with c2:
        end = st.selectbox("종료 월", months, index=len(months) - 1 if months else None, key="f_end")
    with c3:
        buyer = st.selectbox("바이어", ["전체"] + opts['buyers'], key="f_buyer")
    with c4:
        item = st.selectbox("품명", ["전체"] + opts['items'], key="f_item")

    ranking = calc_factory_ranking(raw_rows, cache, start=start, end=end, buyer=buyer, item=item)
    heatmap = calc_region_heatmap(raw_rows, start=start, end=end)

    trend_mark = {'up': '↑ 악화', 'down': '↓ 개선', 'flat': '→ 보합', 'new': '(데이터 부족)'}

    panel_title("🏆 공장별 불량률 랭킹")
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

    panel_title("🗺️ 지역별 불량률 히트맵")
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

    panel_title("💬 코멘트")
    if ranking:
        period = f"{start} ~ {end}" if start and end else "전체기간"
        ai_data = build_ai_comment_data(ranking, period)
        try:
            api_key = st.secrets.get("ANTHROPIC_API_KEY")
        except Exception:
            api_key = None
        if st.button("코멘트 생성", key="comment_btn"):
            with st.spinner("코멘트 생성 중..."):
                comment = get_comment(ai_data, api_key)
            st.markdown(comment)

    panel_title("🔍 공장 상세")
    if ranking:
        factory_names = [r['factory'] for r in ranking]
        selected_factory = st.selectbox("공장 선택", factory_names, key="factory_select")

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
                    key="dl_factory_pdf",
                )
            else:
                st.info("해당 공장의 데이터가 없습니다.")


# ──────────────────────────────────────────────────────────────────
# 탭4: 실적 분석
# ──────────────────────────────────────────────────────────────────
def render_performance_tab():
    panel_title("📈 실적 분석")

    panel_title("데이터 업로드")
    perf_file = st.file_uploader("실적 rawdata 업로드 (Excel)", type=["xlsx"], key="perf_uploader")

    if perf_file is not None and st.button("📥 데이터 로드", type="primary", key="perf_load_btn"):
        with st.spinner("로드 중..."):
            p = tmpdir / perf_file.name
            p.write_bytes(perf_file.getvalue())
            rows = load_performance(str(p))
            st.session_state.perf_rows = rows
        st.success(f"{len(rows):,}건 로드 완료")

    if "perf_rows" not in st.session_state:
        st.info("실적 rawdata 파일을 업로드하고 '데이터 로드' 버튼을 눌러주세요.")
        return

    rows = st.session_state.perf_rows
    opts = get_perf_filter_options(rows)

    panel_title("필터")
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        f_years = st.multiselect("연도", opts['years'], default=opts['years'], key="p_years")
    with c2:
        f_region = st.selectbox("지역", ["전체"] + opts['regions'], key="p_region")
    with c3:
        f_code = st.selectbox("코드", ["전체"] + [f"{c} ({CODE_LABELS.get(c, c)})" for c in opts['codes']], key="p_code")
        f_code_val = f_code.split(" ")[0] if f_code != "전체" else None
    with c4:
        f_buyer = st.selectbox("바이어", ["전체"] + opts['buyers'], key="p_buyer")

    frows = filter_rows(
        rows,
        years=f_years if f_years else None,
        region=None if f_region == "전체" else f_region,
        code=f_code_val,
        buyer=None if f_buyer == "전체" else f_buyer,
    )

    # ── KPI ──────────────────────────────────────────────────────
    panel_title("KPI")
    yearly = summary_by_year(frows)
    if yearly:
        cols = st.columns(len(yearly))
        for col, y in zip(cols, yearly):
            col.metric(f"{y['year']}년 수익", fmt_money(y['revenue']), f"{y['cnt']:,}건")
    else:
        st.info("선택한 조건에 해당하는 데이터가 없습니다.")

    # ── 3개년 월별 추이 ──────────────────────────────────────────
    panel_title("📊 3개년 월별 추이")
    dim_label = st.selectbox("비교 기준", ["전체", "지역별", "바이어별", "브랜드별", "코드별"], key="p_dim")

    dim = None
    group_key = None
    if dim_label == "지역별":
        dim = "region"
        group_key = st.selectbox("지역 선택", opts['regions'], key="p_dim_region")
    elif dim_label == "바이어별":
        dim = "buyer"
        group_key = st.selectbox("바이어 선택", opts['buyers'], key="p_dim_buyer")
    elif dim_label == "브랜드별":
        dim = "brand"
        brand_opts = [b['brand'] for b in summary_by_brand(frows, top_n=30)]
        group_key = st.selectbox("브랜드 선택", brand_opts, key="p_dim_brand")
    elif dim_label == "코드별":
        dim = "code"
        code_opts = {f"{c} ({CODE_LABELS.get(c, c)})": CODE_LABELS.get(c, c) for c in opts['codes']}
        sel = st.selectbox("코드 선택", list(code_opts.keys()), key="p_dim_code")
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
        ymax = pdf['수익'].max() if not pdf.empty else 0
        ticks = [ymax * i / 5 for i in range(6)]
        fig.update_yaxes(tickmode='array', tickvals=ticks,
                          ticktext=[fmt_won_kr(t) for t in ticks], title='수익')
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("표시할 데이터가 없습니다.")

    # ── 동기누적 비교 ─────────────────────────────────────────────
    panel_title("🔁 동기누적 비교")
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
            cum_df_display = cum_df[['연도', '누적건수', '표시']].rename(columns={'표시': '누적수익'}).copy()
            cum_df_display['누적건수'] = cum_df_display['누적건수'].map('{:,}'.format)
            st.dataframe(cum_df_display, use_container_width=True)

            yoy = yoy_comparison(frows, dim=dim, same_months=latest_months, top_n=10, sort_year=latest_year)
            if yoy:
                yoy_df = pd.DataFrame(yoy)
                display_cols = ['label'] + [c for c in yoy_df.columns if c.startswith('y')] + ['growth_24_25', 'growth_25_26']
                display_cols = [c for c in display_cols if c in yoy_df.columns]
                yoy_df_display = yoy_df[display_cols].copy()
                for c in yoy_df_display.columns:
                    if c.startswith('y'):
                        yoy_df_display[c] = yoy_df_display[c].map(fmt_won)
                    elif c.startswith('growth'):
                        yoy_df_display[c] = yoy_df_display[c].map(fmt_pct1)
                st.dataframe(yoy_df_display, use_container_width=True)
        else:
            st.info("최신 연도의 월별 데이터가 없습니다.")

    # ── 목표 vs 실적 (131/152) ───────────────────────────────────
    panel_title("🎯 목표 대비 실적 (131/152)")
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
            empty_actual = {'131': 0, '152': 0, 'total': 0,
                            '131_cnt': 0, '152_cnt': 0, 'total_cnt': 0}
            plan_rows = []
            for mm in [f"{i:02d}" for i in range(1, 13)]:
                p = plan['monthly'].get(mm, {'131': 0, '152': 0, 'total': 0})
                a = actual.get(mm, empty_actual)
                plan_rows.append({
                    '월': mm,
                    '목표(131)': p['131'], '실적(131)': a['131'],
                    '목표(152)': p['152'], '실적(152)': a['152'],
                    '목표(합계)': p['total'], '실적(합계)': a['total'],
                    '달성률(%)': round(a['total'] / p['total'] * 100, 1) if p['total'] else None,
                    '실적건수(131)': a.get('131_cnt', 0),
                    '실적건수(152)': a.get('152_cnt', 0),
                    '실적건수(합계)': a.get('total_cnt', 0),
                })
            plan_df = pd.DataFrame(plan_rows)

            fig2 = go.Figure()
            fig2.add_bar(x=plan_df['월'], y=plan_df['목표(합계)'], name='목표')
            fig2.add_bar(x=plan_df['월'], y=plan_df['실적(합계)'], name='실적')
            fig2.update_layout(barmode='group', title=f"{latest_year}년 월별 목표 대비 실적 (합계)")
            st.plotly_chart(fig2, use_container_width=True)

            money_cols = ['목표(131)', '실적(131)', '목표(152)', '실적(152)', '목표(합계)', '실적(합계)']
            cnt_show_cols = ['실적건수(131)', '실적건수(152)', '실적건수(합계)']
            plan_df_display = plan_df.copy()
            for c in money_cols:
                plan_df_display[c] = plan_df_display[c].map(fmt_won)
            for c in cnt_show_cols:
                plan_df_display[c] = plan_df_display[c].map('{:,}'.format)
            st.dataframe(plan_df_display, use_container_width=True)

            # ── 건수 비교 (전년 대비) ────────────────────────────────
            prev_year = latest_year - 1
            actual_prev = actual_by_month_code(rows, prev_year)
            months_mm = plan_df['월'].tolist()

            st.markdown(f"**월별 실적 건수 비교 (131 / 152, {prev_year}년 vs {latest_year}년)**")
            fig2b = go.Figure()
            fig2b.add_bar(x=months_mm, y=[actual_prev.get(mm, empty_actual)['131_cnt'] for mm in months_mm],
                          name=f'{prev_year} 131(원단검사)')
            fig2b.add_bar(x=months_mm, y=[actual_prev.get(mm, empty_actual)['152_cnt'] for mm in months_mm],
                          name=f'{prev_year} 152(완제품검사)')
            fig2b.add_bar(x=months_mm, y=plan_df['실적건수(131)'], name=f'{latest_year} 131(원단검사)')
            fig2b.add_bar(x=months_mm, y=plan_df['실적건수(152)'], name=f'{latest_year} 152(완제품검사)')
            fig2b.update_layout(barmode='group', title=f"{prev_year}년 vs {latest_year}년 월별 실적 건수")
            st.plotly_chart(fig2b, use_container_width=True)

            cnt_compare_rows = []
            for mm in months_mm:
                cur = actual.get(mm, empty_actual)
                prv = actual_prev.get(mm, empty_actual)
                cnt_compare_rows.append({
                    '월': mm,
                    f'{prev_year}_131': prv.get('131_cnt', 0), f'{latest_year}_131': cur.get('131_cnt', 0),
                    f'{prev_year}_152': prv.get('152_cnt', 0), f'{latest_year}_152': cur.get('152_cnt', 0),
                    f'{prev_year}_합계': prv.get('total_cnt', 0), f'{latest_year}_합계': cur.get('total_cnt', 0),
                })
            cnt_compare_df = pd.DataFrame(cnt_compare_rows)
            cnt_compare_display = cnt_compare_df.copy()
            for c in cnt_compare_display.columns[1:]:
                cnt_compare_display[c] = cnt_compare_display[c].map('{:,}'.format)
            st.dataframe(cnt_compare_display, use_container_width=True)

            # ── 지역별 실적 건수 비교 ──────────────────────────────
            st.markdown("**지역별 월별 실적 건수 비교**")
            cnt_code_label = st.selectbox(
                "코드", ["합계", "131(원단검사)", "152(완제품검사)"], key="cnt_region_code"
            )
            cnt_code = {'합계': None, '131(원단검사)': '131', '152(완제품검사)': '152'}[cnt_code_label]

            region_month_cnt: dict[str, dict[str, int]] = defaultdict(lambda: defaultdict(int))
            for r in rows:
                if r['year'] != latest_year or not r.get('ym') or len(r['ym']) != 7:
                    continue
                if cnt_code and r['code'] != cnt_code:
                    continue
                reg = r['region_label'] or '기타'
                region_month_cnt[reg][r['ym'][-2:]] += 1

            months_all = [f"{i:02d}" for i in range(1, 13)]
            region_order_present = [r for r in REGION_ORDER if r in region_month_cnt] + \
                                    [r for r in region_month_cnt if r not in REGION_ORDER]
            if region_order_present:
                fig2c = go.Figure()
                for reg in region_order_present:
                    fig2c.add_bar(
                        x=months_all,
                        y=[region_month_cnt[reg].get(mm, 0) for mm in months_all],
                        name=reg,
                    )
                fig2c.update_layout(barmode='group', title=f"{latest_year}년 지역별 월별 실적 건수 ({cnt_code_label})")
                st.plotly_chart(fig2c, use_container_width=True)

                region_cnt_df = pd.DataFrame(
                    {'월': months_all,
                     **{reg: [region_month_cnt[reg].get(mm, 0) for mm in months_all] for reg in region_order_present}}
                )
                region_cnt_display = region_cnt_df.copy()
                for reg in region_order_present:
                    region_cnt_display[reg] = region_cnt_display[reg].map('{:,}'.format)
                st.dataframe(region_cnt_display, use_container_width=True)
            else:
                st.info("표시할 데이터가 없습니다.")
        except Exception as e:
            st.warning(f"목표예산 파일을 처리할 수 없습니다: {e}")

    # ── 지역 × 코드 교차분석 ──────────────────────────────────────
    panel_title("🗺️ 지역 × 코드 교차분석")
    ct_year_label = st.selectbox(
        "연도", ["전체"] + [str(y) for y in opts['years']], key="ct_year"
    )
    if ct_year_label == "전체":
        ct_rows = frows
        st.caption(f"기준: 전체 ({', '.join(str(y) for y in opts['years'])}년 합산)")
    else:
        ct_rows = [r for r in frows if r['year'] == int(ct_year_label)]
        st.caption(f"기준: {ct_year_label}년")
    crosstab = region_code_crosstab(ct_rows)
    if crosstab:
        ct_df = pd.DataFrame(crosstab)
        money_df = ct_df[['region', 'c131', 'c152', 'other', 'total']].rename(columns={
            'region': '지역', 'c131': '131(원단검사)', 'c152': '152(완제품검사)',
            'other': '기타', 'total': '합계',
        })
        cnt_df = ct_df[['region', 'cnt131', 'cnt152', 'cnt_other', 'cnt_total']].rename(columns={
            'region': '지역', 'cnt131': '131(원단검사)', 'cnt152': '152(완제품검사)',
            'cnt_other': '기타', 'cnt_total': '합계',
        })

        fig3 = go.Figure()
        fig3.add_bar(x=money_df['지역'], y=money_df['131(원단검사)'], name='131(원단검사)')
        fig3.add_bar(x=money_df['지역'], y=money_df['152(완제품검사)'], name='152(완제품검사)')
        fig3.update_layout(barmode='stack', title="지역별 코드 구성 (금액)")
        ymax3 = money_df['합계'].max() if not money_df.empty else 0
        ticks3 = [ymax3 * i / 5 for i in range(6)]
        fig3.update_yaxes(tickmode='array', tickvals=ticks3,
                           ticktext=[fmt_won_kr(t) for t in ticks3], title='금액')
        st.plotly_chart(fig3, use_container_width=True)

        st.markdown("**금액 (원)**")
        money_df_display = money_df.copy()
        for c in ['131(원단검사)', '152(완제품검사)', '기타', '합계']:
            money_df_display[c] = money_df_display[c].map(fmt_won)
        st.dataframe(money_df_display, use_container_width=True)

        st.markdown("**건수**")
        cnt_df_display = cnt_df.copy()
        for c in ['131(원단검사)', '152(완제품검사)', '기타', '합계']:
            cnt_df_display[c] = cnt_df_display[c].map('{:,}'.format)
        st.dataframe(cnt_df_display, use_container_width=True)
    else:
        st.info("표시할 데이터가 없습니다.")


# ──────────────────────────────────────────────────────────────────
# 상단 탭 구성
# ──────────────────────────────────────────────────────────────────
with st.expander("ℹ️ 사용 안내", expanded=False):
    st.markdown(
        "- **PDF → Excel 변환**: 불량보고서 PDF를 업로드하면 통합 Excel 파일로 변환합니다.\n"
        "- **불량명 표준화**: 불량상세 데이터를 업로드하면 표준 불량명으로 자동 매핑하고, 미매핑/검토 항목을 수동으로 수정할 수 있습니다.\n"
        "- **공장·지역 분석**: (불량명 표준화 탭에서 분석 완료 후) 공장별·지역별 불량률 랭킹, 추이, PDF 보고서를 확인합니다.\n"
        "- **실적 분석**: 실적 rawdata를 업로드하면 월별·브랜드·바이어별 실적을 분석합니다."
    )

tab1, tab2, tab3, tab4 = st.tabs([
    "📄 PDF → Excel",
    "📊 불량명 표준화",
    "🏗️ 공장·지역 분석",
    "📈 실적 분석",
])

def _tab(fn, label):
    try:
        fn()
    except Exception as _e:
        # Streamlit 내부 제어 예외는 이름으로 판단 후 re-raise
        if type(_e).__name__ in ('RerunException', 'StopException', 'ScriptRunnerExitException'):
            raise
        st.error(f"**{label} 오류** — 아래 내용을 개발자에게 전달해주세요:")
        st.exception(_e)

with tab1:
    _tab(render_pdf_tab, "PDF 탭")
with tab2:
    _tab(render_defect_tab, "불량명 탭")
with tab3:
    _tab(render_factory_tab, "공장 탭")
with tab4:
    _tab(render_performance_tab, "실적 탭")
