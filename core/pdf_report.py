"""
pdf_report.py — 공장별 불량률 분석 PDF 보고서 생성
의존성: reportlab, matplotlib
"""
from __future__ import annotations
import io, os
from datetime import datetime
from pathlib import Path

# matplotlib / reportlab 은 함수 호출 시점에 lazy import → 앱 시작 메모리 절감
MATPLOTLIB_OK = None
REPORTLAB_OK  = None

def _ensure_matplotlib():
    global MATPLOTLIB_OK
    if MATPLOTLIB_OK is not None:
        return MATPLOTLIB_OK
    try:
        import matplotlib
        matplotlib.use('Agg')
        import matplotlib.pyplot as _plt
        import matplotlib.font_manager as _fm
        globals()['plt'] = _plt
        globals()['fm']  = _fm
        MATPLOTLIB_OK = True
    except ImportError:
        MATPLOTLIB_OK = False
    return MATPLOTLIB_OK

def _ensure_reportlab():
    global REPORTLAB_OK
    global A4, colors, cm, getSampleStyleSheet, ParagraphStyle
    global TA_CENTER, TA_LEFT, TA_RIGHT
    global SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    global Image, HRFlowable, KeepTogether, pdfmetrics, TTFont
    if REPORTLAB_OK is not None:
        return REPORTLAB_OK
    try:
        from reportlab.lib.pagesizes import A4 as _A4
        from reportlab.lib import colors as _c
        from reportlab.lib.units import cm as _cm
        from reportlab.lib.styles import getSampleStyleSheet as _gss, ParagraphStyle as _PS
        from reportlab.lib.enums import TA_CENTER as _TAC, TA_LEFT as _TAL, TA_RIGHT as _TAR
        from reportlab.platypus import (
            SimpleDocTemplate as _SDT, Paragraph as _P, Spacer as _Sp,
            Table as _T, TableStyle as _TS, Image as _Img,
            HRFlowable as _HR, KeepTogether as _KT,
        )
        from reportlab.pdfbase import pdfmetrics as _pm
        from reportlab.pdfbase.ttfonts import TTFont as _TTF
        globals().update(dict(
            A4=_A4, colors=_c, cm=_cm,
            getSampleStyleSheet=_gss, ParagraphStyle=_PS,
            TA_CENTER=_TAC, TA_LEFT=_TAL, TA_RIGHT=_TAR,
            SimpleDocTemplate=_SDT, Paragraph=_P, Spacer=_Sp,
            Table=_T, TableStyle=_TS, Image=_Img,
            HRFlowable=_HR, KeepTogether=_KT,
            pdfmetrics=_pm, TTFont=_TTF,
        ))
        REPORTLAB_OK = True
    except ImportError:
        REPORTLAB_OK = False
    return REPORTLAB_OK


def _register_korean_font():
    candidates = [
        r"C:\Windows\Fonts\malgun.ttf",
        r"C:\Windows\Fonts\gulim.ttc",
        r"C:\Windows\Fonts\batang.ttc",
        "/usr/share/fonts/truetype/nanum/NanumGothic.ttf",
        "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
    ]
    for path in candidates:
        if os.path.exists(path):
            try:
                pdfmetrics.registerFont(TTFont("Korean", path))
                return "Korean"
            except Exception:
                continue
    return "Helvetica"


def _korean_font_for_matplotlib():
    candidates = [
        r"C:\Windows\Fonts\malgun.ttf",
        r"C:\Windows\Fonts\gulim.ttc",
        r"C:\Windows\Fonts\batang.ttc",
        "/usr/share/fonts/truetype/nanum/NanumGothic.ttf",
    ]
    for path in candidates:
        if os.path.exists(path):
            fm.fontManager.addfont(path)
            return fm.FontProperties(fname=path).get_name()
    return "DejaVu Sans"


def _make_trend_chart(monthly, factory_name, font_name):
    """1차 불량률 단일 추이 차트 (최종 데이터 없을 때 fallback)"""
    if not MATPLOTLIB_OK or not monthly:
        return None
    months = [m["month"] for m in monthly if m.get("rate") is not None]
    rates  = [m["rate"]  for m in monthly if m.get("rate") is not None]
    if not months:
        return None
    plt.rcParams["font.family"] = font_name
    fig, ax = plt.subplots(figsize=(9, 3.5))
    if len(months) == 1:
        ax.bar(months, rates, color='#2B5BA8', width=0.4)
    else:
        ax.plot(months, rates, marker='o', linewidth=2.5, color='#2B5BA8',
                markersize=6, markerfacecolor='white', markeredgewidth=2)
        ax.fill_between(months, rates, alpha=0.08, color='#2B5BA8')
    ax.set_title(f"{factory_name} — 월별 1차 불량률 추이", fontsize=13, pad=12)
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


def _make_dual_trend_chart(monthly, factory_name, font_name):
    """1차 / 최종 불량률 이중 추이 차트"""
    if not MATPLOTLIB_OK or not monthly:
        return None
    valid = [m for m in monthly if m.get("rate") is not None]
    if not valid:
        return None
    months  = [m["month"] for m in valid]
    rates1  = [m["rate"] for m in valid]
    rates2  = [m.get("final_rate") for m in valid]
    corr    = [m.get("correction_rate") for m in valid]
    has_final = any(r is not None and r > 0 for r in rates2)

    plt.rcParams["font.family"] = font_name
    if has_final:
        fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(9, 5.5), sharex=True,
                                        gridspec_kw={'height_ratios': [3, 2]})
    else:
        fig, ax1 = plt.subplots(figsize=(9, 3.5))
        ax2 = None

    # 상단: 1차 / 최종 불량률
    ax1.plot(months, rates1, marker='o', linewidth=2.5, color='#2B5BA8',
             label='1차 불량률', markersize=6, markerfacecolor='white', markeredgewidth=2)
    if has_final:
        _r2 = [r if r is not None else 0 for r in rates2]
        ax1.plot(months, _r2, marker='s', linewidth=2.5, color='#C0392B',
                 label='최종 불량률', markersize=6, markerfacecolor='white', markeredgewidth=2)
    ax1.set_title(f"{factory_name} — 월별 불량률 추이 (1차 vs 최종)", fontsize=12, pad=10)
    ax1.set_ylabel("불량률 (%)")
    ax1.set_ylim(bottom=0)
    ax1.legend(fontsize=9, loc='upper right')
    ax1.grid(axis='y', linestyle='--', alpha=0.4)

    # 하단: 수정합격률
    if ax2 is not None and any(c is not None and c > 0 for c in corr):
        _corr = [c if c is not None else 0 for c in corr]
        _colors = ['#27AE60' if c >= 80 else ('#E67E22' if c >= 60 else '#C0392B')
                   for c in _corr]
        ax2.bar(months, _corr, color=_colors, width=0.5, alpha=0.85)
        for i, (m, c) in enumerate(zip(months, _corr)):
            if c > 0:
                ax2.text(i, c + 1, f'{c:.0f}%', ha='center', va='bottom', fontsize=8)
        ax2.set_ylim(0, 110)
        ax2.set_ylabel("수정합격률 (%)")
        ax2.set_title("수정 합격률 추이", fontsize=10, pad=6)
        ax2.grid(axis='y', linestyle='--', alpha=0.3)
        ax2.axhline(80, color='#27AE60', linewidth=1, linestyle=':', alpha=0.7)

    plt.xticks(rotation=30, ha='right', fontsize=9)
    plt.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=130, bbox_inches='tight')
    plt.close(fig)
    buf.seek(0)
    return buf.read()


def _make_defect_chart(top5, font_name):
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
    ax.set_title(f"주요 불량 유형 TOP {len(top5)}", fontsize=13, pad=12)
    ax.set_xlabel("불량 수량 (개)")
    ax.grid(axis='x', linestyle='--', alpha=0.3)
    plt.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=130, bbox_inches='tight')
    plt.close(fig)
    buf.seek(0)
    return buf.read()


def generate_factory_pdf(detail: dict) -> bytes:
    _ensure_matplotlib()
    if not _ensure_reportlab():
        raise ImportError("reportlab가 설치되지 않았습니다. pip install reportlab")

    font_name = _register_korean_font()
    mpl_font  = _korean_font_for_matplotlib() if MATPLOTLIB_OK else "DejaVu Sans"

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
                            leftMargin=2*cm, rightMargin=2*cm,
                            topMargin=2*cm, bottomMargin=2*cm)

    PRIMARY = colors.HexColor('#2B5BA8')
    LIGHT   = colors.HexColor('#E8F0FE')
    GRAY    = colors.HexColor('#6c757d')
    GREEN   = colors.HexColor('#1e7e34')
    RED     = colors.HexColor('#dc3545')

    getSampleStyleSheet()

    def S(name, **kw):
        kw.setdefault('fontName', font_name)
        kw.setdefault('fontSize', 10)
        kw.setdefault('leading', kw['fontSize'] * 1.4)
        return ParagraphStyle(name, **kw)

    s_title   = S('T2',  fontSize=22, textColor=PRIMARY, alignment=TA_CENTER, spaceAfter=4)
    s_sub     = S('S2',  fontSize=12, textColor=GRAY,    alignment=TA_CENTER, spaceAfter=2)
    s_section = S('Se2', fontSize=13, textColor=PRIMARY, spaceBefore=14, spaceAfter=6)
    s_body    = S('B2',  fontSize=10, leading=16)
    s_cell    = S('C2',  fontSize=9,  leading=13)
    s_cell_c  = S('CC2', fontSize=9,  alignment=TA_CENTER)
    s_cell_r  = S('CR2', fontSize=9,  alignment=TA_RIGHT)

    factory       = detail.get("factory", "")
    region1       = detail.get("region1", "")
    region2       = detail.get("region2", "")
    buyers        = ", ".join(detail.get("buyers", [])) or "—"
    total_inspec  = detail.get("total_inspec", 0)
    total_defect  = detail.get("total_defect", 0)
    total_final   = detail.get("total_final_defect", 0)
    total_second  = detail.get("total_second_inspec", 0)
    record_count  = detail.get("record_count", 0)
    avg_rate      = detail.get("avg_rate") or (round(total_defect / total_inspec * 100, 2) if total_inspec else 0)
    final_rate    = detail.get("final_rate") or 0
    corr_rate     = detail.get("correction_rate") or 0
    has_dual      = total_final > 0 or total_second > 0
    monthly       = detail.get("monthly", [])
    top7          = detail.get("top7_defects", [])
    today         = datetime.now().strftime("%Y년 %m월 %d일")

    story = []

    # ── 헤더: 제목(왼쪽) + FITI CI 로고(오른쪽) ──────────────────
    LOGO_PATH = os.path.join(os.path.dirname(__file__), '..', 'BS 1-06 시그니처_국영문_가로형.png')
    s_title_l = S('TL', fontSize=18, textColor=PRIMARY, alignment=TA_LEFT, spaceAfter=2, leading=24)
    s_fn_l    = S('FL', fontSize=13, textColor=PRIMARY, alignment=TA_LEFT, spaceAfter=2,
                  fontName=font_name, leading=18)
    s_sub_l   = S('SL', fontSize=10, textColor=GRAY,    alignment=TA_LEFT)
    page_w = A4[0] - 3.5 * cm   # 좌우 마진 제외 유효폭
    lw, rw = page_w * 0.58, page_w * 0.42
    left_cells  = [Paragraph("공장 불량률 분석 보고서", s_title_l),
                   Paragraph(factory, s_fn_l),
                   Paragraph(f"보고일: {today}", s_sub_l)]
    right_cells = [Image(LOGO_PATH, width=rw * 0.85, height=rw * 0.85 / 4.54)] \
                  if os.path.exists(LOGO_PATH) else [Paragraph("FITI 시험연구원", s_sub)]
    hdr_tbl = Table([[left_cells, right_cells]], colWidths=[lw, rw])
    hdr_tbl.setStyle(TableStyle([
        ('VALIGN',        (0, 0), (-1, -1), 'MIDDLE'),
        ('ALIGN',         (1, 0), (1,  0),  'RIGHT'),
        ('LINEBELOW',     (0, 0), (-1, -1), 2, PRIMARY),
        ('TOPPADDING',    (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 10),
        ('LEFTPADDING',   (0, 0), (-1, -1), 0),
        ('RIGHTPADDING',  (0, 0), (-1, -1), 0),
    ]))
    story.append(hdr_tbl)
    story.append(Spacer(1, 12))

    # 기본 정보
    story.append(Paragraph("■ 기본 정보", s_section))
    info_data = [
        ["지역", f"{region1} {region2}".strip(), "바이어", buyers],
        ["검사 건수", f"{record_count:,}건", "총 검사수량", f"{total_inspec:,}개"],
    ]
    info_tbl = Table(info_data, colWidths=[3*cm, 6.5*cm, 3*cm, 4.5*cm])
    info_tbl.setStyle(TableStyle([
        ('FONTNAME',   (0,0), (-1,-1), font_name),
        ('FONTSIZE',   (0,0), (-1,-1), 10),
        ('BACKGROUND', (0,0), (0,-1),  LIGHT),
        ('BACKGROUND', (2,0), (2,-1),  LIGHT),
        ('TEXTCOLOR',  (0,0), (0,-1),  PRIMARY),
        ('TEXTCOLOR',  (2,0), (2,-1),  PRIMARY),
        ('GRID',       (0,0), (-1,-1), 0.5, colors.lightgrey),
        ('ROWBACKGROUNDS', (0,0), (-1,-1), [colors.white, colors.HexColor('#f8f9fa')]),
        ('PADDING',    (0,0), (-1,-1), 8),
    ]))
    story.append(info_tbl)
    story.append(Spacer(1, 12))

    # KPI 카드 (1차불량률 / 최종불량률 / 수정합격률 / 총불량수량 / 분석기간)
    rate_color  = GREEN if avg_rate  < 5 else (colors.orange if avg_rate  < 10 else RED)
    final_color = GREEN if final_rate < 5 else (colors.orange if final_rate < 10 else RED)
    corr_color  = GREEN if corr_rate >= 80 else (colors.orange if corr_rate >= 60 else RED)
    if has_dual:
        kpi_data = [[
            Paragraph(f"<b>{avg_rate:.2f}%</b><br/><font size=8 color='grey'>1차 불량률</font>", s_cell_c),
            Paragraph(f"<b>{final_rate:.2f}%</b><br/><font size=8 color='grey'>최종 불량률</font>", s_cell_c),
            Paragraph(f"<b>{corr_rate:.1f}%</b><br/><font size=8 color='grey'>수정 합격률</font>", s_cell_c),
            Paragraph(f"<b>{total_defect:,}</b><br/><font size=8 color='grey'>1차불량수량</font>", s_cell_c),
            Paragraph(f"<b>{len(monthly)}</b><br/><font size=8 color='grey'>분석 기간(월)</font>", s_cell_c),
        ]]
        kpi_tbl = Table(kpi_data, colWidths=[3.3*cm]*5)
        kpi_tbl.setStyle(TableStyle([
            ('FONTNAME',   (0,0), (-1,-1), font_name),
            ('FONTSIZE',   (0,0), (-1,-1), 15),
            ('TEXTCOLOR',  (0,0), (0,0),   rate_color),
            ('TEXTCOLOR',  (1,0), (1,0),   final_color),
            ('TEXTCOLOR',  (2,0), (2,0),   corr_color),
            ('TEXTCOLOR',  (3,0), (3,0),   PRIMARY),
            ('TEXTCOLOR',  (4,0), (4,0),   GRAY),
            ('ALIGN',      (0,0), (-1,-1), 'CENTER'),
            ('VALIGN',     (0,0), (-1,-1), 'MIDDLE'),
            ('BOX',        (0,0), (0,0),   1, colors.lightgrey),
            ('BOX',        (1,0), (1,0),   1, colors.lightgrey),
            ('BOX',        (2,0), (2,0),   1, colors.lightgrey),
            ('BOX',        (3,0), (3,0),   1, colors.lightgrey),
            ('BOX',        (4,0), (4,0),   1, colors.lightgrey),
            ('ROWBACKGROUNDS', (0,0), (-1,-1), [LIGHT]),
            ('PADDING',    (0,0), (-1,-1), 12),
        ]))
    else:
        kpi_data = [[
            Paragraph(f"<b>{avg_rate:.2f}%</b><br/><font size=9 color='grey'>평균 불량률</font>", s_cell_c),
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

    # 월별 추이 차트 (1차+최종 데이터가 있으면 이중, 없으면 단일)
    story.append(Paragraph("■ 월별 불량률 추이", s_section))
    if has_dual:
        trend_png = _make_dual_trend_chart(monthly, factory, mpl_font)
        chart_h = 9.5 * cm
    else:
        trend_png = _make_trend_chart(monthly, factory, mpl_font)
        chart_h = 6.5 * cm
    if trend_png:
        story.append(Image(io.BytesIO(trend_png), width=16*cm, height=chart_h))
    else:
        story.append(Paragraph("(차트 생성 불가 — matplotlib 설치 필요)", s_body))
    story.append(Spacer(1, 12))

    # 불량 TOP5
    if top7:
        story.append(Paragraph(f"■ 주요 불량 유형 TOP {len(top7)}", s_section))
        defect_png = _make_defect_chart(top7, mpl_font)
        if defect_png:
            story.append(Image(io.BytesIO(defect_png), width=16*cm, height=5.5*cm))
        d_rows = [[
            Paragraph("<b>순위</b>", s_cell_c),
            Paragraph("<b>표준 불량명</b>", s_cell_c),
            Paragraph("<b>불량 수량</b>", s_cell_c),
            Paragraph("<b>비율</b>", s_cell_c),
        ]] + [
            [Paragraph(f"{i+1}위", s_cell_c),
             Paragraph(d["name"], s_cell),
             Paragraph(f"{d['qty']:,}개", s_cell_r),
             Paragraph(f"{d['pct']}%", s_cell_c)]
            for i, d in enumerate(top7)
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

    # 월별 상세 테이블
    story.append(Paragraph("■ 월별 검사 실적", s_section))
    if has_dual:
        m_rows = [[
            Paragraph("<b>연월</b>", s_cell_c),
            Paragraph("<b>검사수량</b>", s_cell_c),
            Paragraph("<b>1차불량수량</b>", s_cell_c),
            Paragraph("<b>1차불량률</b>", s_cell_c),
            Paragraph("<b>최종불량수량</b>", s_cell_c),
            Paragraph("<b>최종불량률</b>", s_cell_c),
            Paragraph("<b>수정합격률</b>", s_cell_c),
        ]]
        for m in monthly:
            r1 = m.get("rate");  r2 = m.get("final_rate"); rc = m.get("correction_rate")
            m_rows.append([
                Paragraph(m["month"], s_cell_c),
                Paragraph(f"{m.get('inspec',0):,}", s_cell_r),
                Paragraph(f"{m.get('defect',0):,}", s_cell_r),
                Paragraph(f"{r1:.2f}%" if r1 is not None else "—", s_cell_c),
                Paragraph(f"{m.get('final_defect',0):,}", s_cell_r),
                Paragraph(f"{r2:.2f}%" if r2 is not None else "—", s_cell_c),
                Paragraph(f"{rc:.1f}%" if rc is not None else "—", s_cell_c),
            ])
        m_tbl = Table(m_rows, colWidths=[2.2*cm, 2.4*cm, 2.4*cm, 2.2*cm, 2.4*cm, 2.2*cm, 2.2*cm])
    else:
        m_rows = [[
            Paragraph("<b>연월</b>", s_cell_c),
            Paragraph("<b>검사수량</b>", s_cell_c),
            Paragraph("<b>불량수량</b>", s_cell_c),
            Paragraph("<b>불량률</b>", s_cell_c),
            Paragraph("<b>평가</b>", s_cell_c),
        ]]
        for m in monthly:
            rate = m.get("rate")
            if rate is None:
                rate_str, eval_str = "—", "—"
            else:
                rate_str = f"{rate:.2f}%"
                if rate < 1.0:   eval_str = "우수"
                elif rate < 2.0: eval_str = "양호"
                elif rate < 3.5: eval_str = "주의"
                else:            eval_str = "불량"
            m_rows.append([
                Paragraph(m["month"], s_cell_c),
                Paragraph(f"{m.get('inspec',0):,}", s_cell_r),
                Paragraph(f"{m.get('defect',0):,}", s_cell_r),
                Paragraph(rate_str, s_cell_c),
                Paragraph(eval_str, s_cell_c),
            ])
        m_tbl = Table(m_rows, colWidths=[3*cm, 3.5*cm, 3.5*cm, 3*cm, 3.5*cm])
    m_tbl.setStyle(TableStyle([
        ('FONTNAME',   (0,0), (-1,-1), font_name),
        ('BACKGROUND', (0,0), (-1,0),  PRIMARY),
        ('TEXTCOLOR',  (0,0), (-1,0),  colors.white),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#f8f9fa')]),
        ('GRID',       (0,0), (-1,-1), 0.5, colors.lightgrey),
        ('PADDING',    (0,0), (-1,-1), 7),
        ('ALIGN',      (0,0), (-1,-1), 'CENTER'),
    ]))
    story.append(m_tbl)
    story.append(Spacer(1, 20))

    # 푸터
    story.append(HRFlowable(width="100%", thickness=1, color=colors.lightgrey, spaceAfter=6))
    story.append(Paragraph(
        f"본 보고서는 FITI 제품평가팀 검사 데이터 기반으로 자동 생성되었습니다. | {today}",
        S('Footer', fontSize=8, textColor=GRAY, alignment=TA_CENTER, fontName=font_name)))

    doc.build(story)
    buf.seek(0)
    return buf.read()
