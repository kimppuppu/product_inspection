"""
pdf_report.py — 공장별 불량률 분석 PDF 보고서 생성
의존성: reportlab, matplotlib
"""
from __future__ import annotations
import io, os
from datetime import datetime
from pathlib import Path

try:
    import matplotlib
    matplotlib.use('Agg')  # GUI 없이 백엔드 렌더링
    import matplotlib.pyplot as plt
    import matplotlib.font_manager as fm
    MATPLOTLIB_OK = True
except ImportError:
    MATPLOTLIB_OK = False

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.lib.units import cm
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
    from reportlab.platypus import (
        SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
        Image, HRFlowable, KeepTogether
    )
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    REPORTLAB_OK = True
except ImportError:
    REPORTLAB_OK = False


# ── 한글 폰트 등록 ───────────────────────────────────────────────
def _register_korean_font():
    """Windows 시스템 말굴림 폰트 또는 대체 폰트 등록"""
    candidates = [
        r"C:\Windows\Fonts\malgun.ttf",       # 맑은 고딕
        r"C:\Windows\Fonts\gulim.ttc",         # 굴림
        r"C:\Windows\Fonts\batang.ttc",        # 바탕
        "/usr/share/fonts/truetype/nanum/NanumGothic.ttf",  # Linux
        "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
    ]
    for path in candidates:
        if os.path.exists(path):
            try:
                pdfmetrics.registerFont(TTFont("Korean", path))
                return "Korean"
            except Exception:
                continue
    return "Helvetica"  # 폴백 (한글 깨질 수 있음)


def _korean_font_for_matplotlib():
    """matplotlib용 한글 폰트 이름 반환 — 실제 등록된 이름 사용"""
    candidates = [
        r"C:\Windows\Fonts\malgun.ttf",
        r"C:\Windows\Fonts\gulim.ttc",
        r"C:\Windows\Fonts\batang.ttc",
        "/usr/share/fonts/truetype/nanum/NanumGothic.ttf",
    ]
    for path in candidates:
        if os.path.exists(path):
            fm.fontManager.addfont(path)
            font_name = fm.FontProperties(fname=path).get_name()
            return font_name
    return "DejaVu Sans"


# ── 차트 생성 ────────────────────────────────────────────────────
def _make_trend_chart(monthly: list[dict], factory_name: str, font_name: str) -> bytes | None:
    """월별 불량률 추이 선 그래프 → PNG bytes"""
    if not MATPLOTLIB_OK or not monthly:
        return None
    months = [m["month"] for m in monthly if m.get("rate") is not None]
    rates  = [m["rate"]  for m in monthly if m.get("rate") is not None]
    if len(months) < 1:
        return None

    plt.rcParams["font.family"] = font_name
    fig, ax = plt.subplots(figsize=(9, 3.5))
    if len(months) == 1:
        ax.bar(months, rates, color='#2B5BA8', width=0.4)
    else:
        ax.plot(months, rates, marker='o', linewidth=2.5, color='#2B5BA8',
                markersize=6, markerfacecolor='white', markeredgewidth=2)
        ax.fill_between(months, rates, alpha=0.08, color='#2B5BA8')
    ax.set_title(f"{factory_name} — 월별 불량률 추이", fontsize=13, pad=12)
    ax.set_ylabel("불량률 (%)")
    ax.set_ylim(bottom=0)
    ax.grid(axis='y', linestyle='--', alpha=0.4)
    plt.xticks(rotation=30, ha='right', fontsize=9)
    plt.tight_layout()

    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=130, bbox_inches='tight')
    plt.close(fig)
    buf.seek(0)
    return buf.read()


def _make_defect_chart(top5: list[dict], font_name: str) -> bytes | None:
    """불량 TOP5 수평 막대 그래프 → PNG bytes"""
    if not MATPLOTLIB_OK or not top5:
        return None

    names = [d["name"] for d in reversed(top5)]
    qtys  = [d["qty"]  for d in reversed(top5)]
    bar_colors = ['#2B5BA8','#4472C4','#5B9BD5','#70AD47','#ED7D31'][::-1]

    plt.rcParams["font.family"] = font_name
    fig, ax = plt.subplots(figsize=(9, 3))
    bars = ax.barh(names, qtys, color=bar_colors[:len(names)], edgecolor='white')
    for bar, qty in zip(bars, qtys):
        ax.text(bar.get_width() + max(qtys)*0.01, bar.get_y() + bar.get_height()/2,
                f'{qty:,}개', va='center', fontsize=9)
    ax.set_title("주요 불량 유형 TOP 5", fontsize=13, pad=12)
    ax.set_xlabel("불량 수량 (개)")
    ax.grid(axis='x', linestyle='--', alpha=0.3)
    plt.tight_layout()

    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=130, bbox_inches='tight')
    plt.close(fig)
    buf.seek(0)
    return buf.read()


# ── PDF 메인 생성 함수 ───────────────────────────────────────────
def generate_factory_pdf(detail: dict) -> bytes:
    """
    detail: calc_factory_detail() 반환값
    반환: PDF bytes
    """
    if not REPORTLAB_OK:
        raise ImportError("reportlab가 설치되지 않았습니다. pip install reportlab")

    font_name = _register_korean_font()
    mpl_font  = _korean_font_for_matplotlib() if MATPLOTLIB_OK else "DejaVu Sans"

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=2*cm, rightMargin=2*cm,
        topMargin=2*cm, bottomMargin=2*cm
    )

    # 스타일
    PRIMARY = colors.HexColor('#2B5BA8')
    LIGHT   = colors.HexColor('#E8F0FE')
    GRAY    = colors.HexColor('#6c757d')
    GREEN   = colors.HexColor('#1e7e34')
    RED     = colors.HexColor('#dc3545')

    styles = getSampleStyleSheet()
    def S(name, **kw):
        kw.pop('parent', None)          # parent 무시 — fontName 충돌 방지
        kw.setdefault('fontName', font_name)
        kw.setdefault('fontSize', 10)
        kw.setdefault('leading',  kw['fontSize'] * 1.4)
        return ParagraphStyle(name, **kw)

    s_title   = S('Title2',   fontSize=22, textColor=PRIMARY,
                  alignment=TA_CENTER, spaceAfter=4)
    s_sub     = S('Sub2',     fontSize=12, textColor=GRAY,
                  alignment=TA_CENTER, spaceAfter=2)
    s_section = S('Section2', fontSize=13, textColor=PRIMARY,
                  spaceBefore=14, spaceAfter=6)
    s_body    = S('Body2',    fontSize=10, leading=16)
    s_cell    = S('Cell2',    fontSize=9,  leading=13)
    s_cell_c  = S('CellC2',  fontSize=9,  alignment=TA_CENTER)
    s_cell_r  = S('CellR2',  fontSize=9,  alignment=TA_RIGHT)

    factory = detail.get("factory", "")
    region1 = detail.get("region1", "")
    region2 = detail.get("region2", "")
    buyers  = ", ".join(detail.get("buyers", [])) or "—"
    total_inspec = detail.get("total_inspec", 0)
    total_defect = detail.get("total_defect", 0)
    record_count = detail.get("record_count", 0)
    avg_rate = round(total_defect / total_inspec * 100, 2) if total_inspec else 0
    monthly  = detail.get("monthly", [])
    top5     = detail.get("top5_defects", [])
    today    = datetime.now().strftime("%Y년 %m월 %d일")

    story = []

    # ── 헤더 ──────────────────────────────────────────────────────
    story.append(HRFlowable(width="100%", thickness=4, color=PRIMARY, spaceAfter=12))
    story.append(Paragraph(f"공장 불량률 분석 보고서", s_title))
    story.append(Paragraph(factory, S('FName', fontSize=18, textColor=colors.black,
                                      alignment=TA_CENTER, spaceAfter=4,
                                      fontName=font_name)))
    story.append(Paragraph(f"보고일: {today}", s_sub))
    story.append(HRFlowable(width="100%", thickness=1, color=colors.lightgrey, spaceAfter=16))

    # ── 기본 정보 카드 ──────────────────────────────────────────────
    story.append(Paragraph("■ 기본 정보", s_section))
    info_data = [
        ["지역", f"{region1} {region2}".strip(), "바이어", buyers],
        ["검사 건수", f"{record_count:,}건", "총 검사수량", f"{total_inspec:,}개"],
    ]
    info_tbl = Table(info_data, colWidths=[3*cm, 6.5*cm, 3*cm, 4.5*cm])
    info_tbl.setStyle(TableStyle([
        ('FONTNAME',   (0,0), (-1,-1), font_name),
        ('FONTSIZE',   (0,0), (-1,-1), 10),
        ('BACKGROUND', (0,0), (0,-1), LIGHT),
        ('BACKGROUND', (2,0), (2,-1), LIGHT),
        ('FONTNAME',   (0,0), (0,-1), font_name),
        ('FONTNAME',   (2,0), (2,-1), font_name),
        ('TEXTCOLOR',  (0,0), (0,-1), PRIMARY),
        ('TEXTCOLOR',  (2,0), (2,-1), PRIMARY),
        ('GRID',       (0,0), (-1,-1), 0.5, colors.lightgrey),
        ('ROWBACKGROUNDS', (0,0), (-1,-1), [colors.white, colors.HexColor('#f8f9fa')]),
        ('PADDING',    (0,0), (-1,-1), 8),
    ]))
    story.append(info_tbl)
    story.append(Spacer(1, 12))

    # ── 핵심 지표 3개 카드 ──────────────────────────────────────────
    rate_color = GREEN if avg_rate < 1.5 else (colors.orange if avg_rate < 3 else RED)
    kpi_data = [[
        Paragraph(f"<b>{avg_rate:.1f}%</b><br/><font size=9 color='grey'>평균 불량률</font>", s_cell_c),
        Paragraph(f"<b>{total_defect:,}</b><br/><font size=9 color='grey'>총 불량수량</font>", s_cell_c),
        Paragraph(f"<b>{len(monthly)}</b><br/><font size=9 color='grey'>분석 기간(월)</font>", s_cell_c),
    ]]
    kpi_tbl = Table(kpi_data, colWidths=[5.5*cm]*3)
    kpi_tbl.setStyle(TableStyle([
        ('FONTNAME',   (0,0), (-1,-1), font_name),
        ('FONTSIZE',   (0,0), (-1,-1), 16),
        ('TEXTCOLOR',  (0,0), (0,0),   rate_color),
        ('TEXTCOLOR',  (1,0), (1,0),   PRIMARY),
        ('TEXTCOLOR',  (2,0), (2,0),   GRAY),
        ('ALIGN',      (0,0), (-1,-1), 'CENTER'),
        ('VALIGN',     (0,0), (-1,-1), 'MIDDLE'),
        ('BOX',        (0,0), (0,0),   1, colors.lightgrey),
        ('BOX',        (1,0), (1,0),   1, colors.lightgrey),
        ('BOX',        (2,0), (2,0),   1, colors.lightgrey),
        ('ROWBACKGROUNDS', (0,0), (-1,-1), [LIGHT]),
        ('PADDING',    (0,0), (-1,-1), 14),
    ]))
    story.append(kpi_tbl)
    story.append(Spacer(1, 16))

    # ── 월별 추이 차트 ──────────────────────────────────────────────
    story.append(Paragraph("■ 월별 불량률 추이", s_section))
    trend_png = _make_trend_chart(monthly, factory, mpl_font)
    if trend_png:
        story.append(Image(io.BytesIO(trend_png), width=16*cm, height=6.5*cm))
    else:
        story.append(Paragraph("(차트 생성 불가 — matplotlib 설치 필요)", s_body))
    story.append(Spacer(1, 12))

    # ── 불량 TOP5 차트 ──────────────────────────────────────────────
    if top5:
        story.append(Paragraph("■ 주요 불량 유형 TOP 5", s_section))
        defect_png = _make_defect_chart(top5, mpl_font)
        if defect_png:
            story.append(Image(io.BytesIO(defect_png), width=16*cm, height=5.5*cm))

        # 테이블
        d_header = [
            Paragraph("<b>순위</b>", s_cell_c),
            Paragraph("<b>표준 불량명</b>", s_cell_c),
            Paragraph("<b>불량 수량</b>", s_cell_c),
            Paragraph("<b>비율</b>", s_cell_c),
        ]
        d_rows = [d_header] + [
            [Paragraph(f"{i+1}위", s_cell_c),
             Paragraph(d["name"], s_cell),
             Paragraph(f"{d['qty']:,}개", s_cell_r),
             Paragraph(f"{d['pct']}%", s_cell_c)]
            for i, d in enumerate(top5)
        ]
        d_tbl = Table(d_rows, colWidths=[2*cm, 9*cm, 3*cm, 2.5*cm])
        d_tbl.setStyle(TableStyle([
            ('FONTNAME',   (0,0), (-1,-1), font_name),
            ('BACKGROUND', (0,0), (-1,0),  PRIMARY),
            ('TEXTCOLOR',  (0,0), (-1,0),  colors.white),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#f8f9fa')]),
            ('GRID',       (0,0), (-1,-1), 0.5, colors.lightgrey),
            ('PADDING',    (0,0), (-1,-1), 7),
            ('ALIGN',      (2,0), (-1,-1), 'CENTER'),
        ]))
        story.append(Spacer(1, 8))
        story.append(d_tbl)
        story.append(Spacer(1, 12))

    # ── 월별 상세 테이블 ────────────────────────────────────────────
    story.append(Paragraph("■ 월별 검사 실적", s_section))
    m_header = [
        Paragraph("<b>연월</b>", s_cell_c),
        Paragraph("<b>검사수량</b>", s_cell_c),
        Paragraph("<b>불량수량</b>", s_cell_c),
        Paragraph("<b>불량률</b>", s_cell_c),
        Paragraph("<b>평가</b>", s_cell_c),
    ]
    m_rows = [m_header]
    for m in monthly:
        rate = m.get("rate")
        if rate is None:
            rate_str, eval_str, rc = "—", "—", colors.white
        else:
            rate_str = f"{rate:.2f}%"
            if rate < 1.0:   eval_str, rc = "✅ 우수", colors.HexColor('#C6EFCE')
            elif rate < 2.0: eval_str, rc = "양호",  colors.white
            elif rate < 3.5: eval_str, rc = "⚠️ 주의", colors.HexColor('#FFEB9C')
            else:            eval_str, rc = "❌ 불량", colors.HexColor('#FFC7CE')
        m_rows.append([
            Paragraph(m["month"], s_cell_c),
            Paragraph(f"{m.get('inspec',0):,}", s_cell_r),
            Paragraph(f"{m.get('defect',0):,}", s_cell_r),
            Paragraph(rate_str, s_cell_c),
            Paragraph(eval_str, s_cell_c),
        ])

    m_tbl = Table(m_rows, colWidths=[3*cm, 3.5*cm, 3.5*cm, 3*cm, 3.5*cm])
    # 평가 열 색상
    style_cmds = [
        ('FONTNAME',   (0,0), (-1,-1), font_name),
        ('BACKGROUND', (0,0), (-1,0),  PRIMARY),
        ('TEXTCOLOR',  (0,0), (-1,0),  colors.white),
        ('GRID',       (0,0), (-1,-1), 0.5, colors.lightgrey),
        ('PADDING',    (0,0), (-1,-1), 7),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#f8f9fa')]),
    ]
    for i, m in enumerate(monthly, 1):
        rate = m.get("rate")
        if rate is not None:
            if rate < 1.0:   bg = colors.HexColor('#C6EFCE')
            elif rate < 2.0: bg = colors.white
            elif rate < 3.5: bg = colors.HexColor('#FFEB9C')
            else:            bg = colors.HexColor('#FFC7CE')
            style_cmds.append(('BACKGROUND', (4,i), (4,i), bg))
    m_tbl.setStyle(TableStyle(style_cmds))
    story.append(m_tbl)

    # ── 푸터 ──────────────────────────────────────────────────────
    story.append(Spacer(1, 20))
    story.append(HRFlowable(width="100%", thickness=1, color=colors.lightgrey))
    story.append(Paragraph(
        f"<font color='grey' size=8>본 보고서는 불량보고서 분석 시스템에서 자동 생성되었습니다. | {today}</font>",
        S('Foooter', fontSize=8, alignment=TA_CENTER, textColor=GRAY, spaceBefore=6)
    ))

    doc.build(story)
    buf.seek(0)
    return buf.read()
