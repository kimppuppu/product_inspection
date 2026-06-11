"""
1_PDF_변환.py — 불량보고서 PDF → Excel 변환
"""
import sys, tempfile
from pathlib import Path
from datetime import datetime

import streamlit as st

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from core.pdf_extractor import parse_pdf, make_workbook

st.set_page_config(page_title="PDF → Excel 변환", page_icon="📄", layout="wide")
st.title("📄 PDF → Excel 변환")

st.markdown("불량보고서 PDF 파일들을 업로드하면 통합 Excel 파일로 변환합니다.")

uploaded_files = st.file_uploader(
    "PDF 파일 업로드 (여러 개 선택 가능)",
    type=["pdf"],
    accept_multiple_files=True,
)

if uploaded_files:
    st.write(f"업로드된 파일: {len(uploaded_files)}개")

    if st.button("🚀 변환 시작", type="primary"):
        records = []
        failed = []
        progress = st.progress(0.0)
        status = st.empty()
        log_box = st.container(height=250)

        with tempfile.TemporaryDirectory() as tmpdir:
            total = len(uploaded_files)
            for i, uf in enumerate(sorted(uploaded_files, key=lambda f: f.name), 1):
                status.write(f"({i}/{total}) 변환 중: {uf.name}")
                tmp_path = Path(tmpdir) / uf.name
                tmp_path.write_bytes(uf.getvalue())
                try:
                    rec = parse_pdf(str(tmp_path))
                    records.append(rec)
                    log_box.write(f"✅ 완료: {rec.get('REPORT NO.', '')} / {rec.get('공장', '')}")
                except Exception as e:
                    failed.append(uf.name)
                    log_box.write(f"❌ 실패: {uf.name} — {e}")
                progress.progress(i / total)

            if not records:
                st.error("변환에 성공한 PDF가 없습니다.")
            else:
                ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                out_name = f"불량율_분석_통합_{ts}.xlsx"
                out_path = Path(tmpdir) / out_name
                make_workbook(records, str(out_path))

                st.success(f"✅ 완료 — 성공 {len(records)}개 / 실패 {len(failed)}개")
                if failed:
                    st.warning("실패한 파일: " + ", ".join(failed))

                st.download_button(
                    "⬇️ 통합 Excel 다운로드",
                    data=out_path.read_bytes(),
                    file_name=out_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                )
else:
    st.info("PDF 파일을 업로드해주세요.")
