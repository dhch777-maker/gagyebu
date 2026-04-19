# -*- coding: utf-8 -*-
"""
이마트 에브리데이 사업제안서 PPT 리브랜딩 빌더
- 원본 데이터(차트/테이블) 유지, 브랜드 톤만 전면 리스킨
- 결과: 이마트에브리데이_사업제안서_v2_260301.pptx
"""
import sys
import os
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
from pptx.oxml.ns import qn
from lxml import etree

sys.stdout.reconfigure(encoding='utf-8')

# ============================================================
# 브랜드 컬러 팔레트
# ============================================================
RED = RGBColor(0xE6, 0x00, 0x12)
RED_SOFT = RGBColor(0xFF, 0xE5, 0xE8)
YELLOW = RGBColor(0xFF, 0xD2, 0x00)
YELLOW_SOFT = RGBColor(0xFF, 0xF6, 0xCC)
DARK = RGBColor(0x1A, 0x1A, 0x1A)
MID = RGBColor(0x66, 0x66, 0x66)
LIGHT = RGBColor(0xF5, 0xF5, 0xF5)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BORDER = RGBColor(0xE0, 0xE0, 0xE0)

FONT_TITLE = "Pretendard"
FONT_BODY = "Pretendard"
FONT_FALLBACK = "맑은 고딕"

# ============================================================
# Presentation 초기화 (16:9)
# ============================================================
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
SW = prs.slide_width
SH = prs.slide_height

BLANK = prs.slide_layouts[6]


# ============================================================
# 공통 헬퍼
# ============================================================
def add_rect(slide, x, y, w, h, fill=None, line=None, line_w=None):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    shp.shadow.inherit = False
    if fill is not None:
        shp.fill.solid()
        shp.fill.fore_color.rgb = fill
    else:
        shp.fill.background()
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
        if line_w is not None:
            shp.line.width = line_w
    return shp


def add_rounded(slide, x, y, w, h, fill=None, line=None, radius=0.05):
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    shp.shadow.inherit = False
    shp.adjustments[0] = radius
    if fill is not None:
        shp.fill.solid()
        shp.fill.fore_color.rgb = fill
    else:
        shp.fill.background()
    if line is None:
        shp.line.fill.background()
    else:
        shp.line.color.rgb = line
        shp.line.width = Pt(1)
    return shp


def add_text(slide, x, y, w, h, text, *, size=14, bold=False, color=DARK,
             font=FONT_BODY, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
             line_spacing=1.2):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.margin_left = Pt(0)
    tf.margin_right = Pt(0)
    tf.margin_top = Pt(0)
    tf.margin_bottom = Pt(0)
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    lines = text.split("\n")
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = line_spacing
        r = p.add_run()
        r.text = line
        r.font.name = font
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.color.rgb = color
    return tb


def add_left_bar(slide):
    """좌측 레드 컬러바"""
    add_rect(slide, 0, 0, Inches(0.08), SH, fill=RED)


def add_page_header(slide, page_num, total=14):
    """우상단 페이지 번호 + 상단 얇은 라인"""
    add_rect(slide, Inches(0.6), Inches(0.45), Inches(0.3), Inches(0.03), fill=RED)
    add_text(
        slide, Inches(0.95), Inches(0.35), Inches(5), Inches(0.3),
        "EMART EVERYDAY  ·  사업제안",
        size=10, bold=True, color=MID, font=FONT_TITLE,
    )
    add_text(
        slide, SW - Inches(1.6), Inches(0.35), Inches(1.0), Inches(0.3),
        f"{page_num:02d} / {total:02d}",
        size=10, bold=True, color=RED, font=FONT_TITLE, align=PP_ALIGN.RIGHT,
    )


def add_headline(slide, headline, subhead=None, y=Inches(0.9)):
    """상단 헤드라인 + 서브카피"""
    add_text(
        slide, Inches(0.6), y, Inches(12), Inches(0.8),
        headline, size=32, bold=True, color=DARK, font=FONT_TITLE,
    )
    if subhead:
        add_text(
            slide, Inches(0.62), y + Inches(0.9), Inches(12), Inches(0.4),
            subhead, size=14, color=MID, font=FONT_BODY,
        )
    # 헤드라인 아래 얇은 레드 라인
    add_rect(slide, Inches(0.6), y + Inches(0.78), Inches(0.5), Inches(0.04), fill=RED)


def add_footer(slide, text="이마트 에브리데이 사업제안서  ·  2026.03.01"):
    add_text(
        slide, Inches(0.6), SH - Inches(0.4), Inches(12), Inches(0.3),
        text, size=9, color=MID, font=FONT_BODY,
    )


def add_big_number(slide, x, y, w, h, number, unit="", color=RED):
    """초대형 숫자 강조 박스"""
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.margin_left = Pt(0)
    tf.margin_right = Pt(0)
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = number
    r.font.name = FONT_TITLE
    r.font.size = Pt(64)
    r.font.bold = True
    r.font.color.rgb = color
    if unit:
        r2 = p.add_run()
        r2.text = unit
        r2.font.name = FONT_TITLE
        r2.font.size = Pt(20)
        r2.font.bold = True
        r2.font.color.rgb = color


def add_icon_circle(slide, cx, cy, r, color=RED, symbol=""):
    """심볼이 들어간 원형 아이콘"""
    shp = slide.shapes.add_shape(MSO_SHAPE.OVAL, cx - r, cy - r, r * 2, r * 2)
    shp.shadow.inherit = False
    shp.fill.solid()
    shp.fill.fore_color.rgb = color
    shp.line.fill.background()
    if symbol:
        tb = slide.shapes.add_textbox(cx - r, cy - r, r * 2, r * 2)
        tf = tb.text_frame
        tf.margin_left = Pt(0)
        tf.margin_right = Pt(0)
        tf.margin_top = Pt(0)
        tf.margin_bottom = Pt(0)
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = symbol
        run.font.name = FONT_TITLE
        run.font.size = Pt(int(r / Emu(1) * 914400 * 0 + 24))
        run.font.bold = True
        run.font.color.rgb = WHITE


# ============================================================
# 아이콘 박스(심볼형)
# ============================================================
def icon_symbol_box(slide, x, y, size, symbol, color=RED, bg=WHITE, border=RED):
    """컬러 원 + Unicode 심볼"""
    circle = slide.shapes.add_shape(MSO_SHAPE.OVAL, x, y, size, size)
    circle.shadow.inherit = False
    circle.fill.solid()
    circle.fill.fore_color.rgb = bg
    circle.line.color.rgb = border
    circle.line.width = Pt(2.5)

    tb = slide.shapes.add_textbox(x, y, size, size)
    tf = tb.text_frame
    tf.margin_left = Pt(0)
    tf.margin_right = Pt(0)
    tf.margin_top = Pt(0)
    tf.margin_bottom = Pt(0)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = symbol
    r.font.name = FONT_TITLE
    r.font.size = Pt(int(Emu(size).inches * 18))
    r.font.bold = True
    r.font.color.rgb = color


def add_arrow(slide, x, y, w, h, color=RED):
    shp = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, x, y, w, h)
    shp.shadow.inherit = False
    shp.fill.solid()
    shp.fill.fore_color.rgb = color
    shp.line.fill.background()


# ============================================================
# Slide 1 — 표지: 매일, 동네가 모이는 거점
# ============================================================
def slide_01():
    s = prs.slides.add_slide(BLANK)
    # 배경: 상단 2/3 레드, 하단 1/3 화이트
    add_rect(s, 0, 0, SW, SH, fill=WHITE)
    add_rect(s, 0, 0, SW, Inches(4.8), fill=RED)

    # 좌측 옐로우 대각선 포인트
    add_rect(s, 0, Inches(4.6), Inches(1.5), Inches(0.1), fill=YELLOW)

    # 로고 느낌 상단
    add_text(
        s, Inches(0.8), Inches(0.7), Inches(8), Inches(0.4),
        "EMART EVERYDAY  ×  사업제안",
        size=14, bold=True, color=YELLOW, font=FONT_TITLE,
    )

    # 메인 카피
    add_text(
        s, Inches(0.8), Inches(1.8), Inches(11.5), Inches(2.0),
        "매일, 동네가 모이는 거점.",
        size=60, bold=True, color=WHITE, font=FONT_TITLE, line_spacing=1.05,
    )
    # 서브 카피
    add_text(
        s, Inches(0.8), Inches(3.3), Inches(11.5), Inches(1.0),
        "이마트 에브리데이 점포를 지역 공동구매 수령 거점으로.\n고정비 0원, 방문 유입 상시, 수익은 자동.",
        size=18, color=WHITE, font=FONT_BODY, line_spacing=1.4,
    )

    # 하단 블록
    add_text(
        s, Inches(0.8), Inches(5.2), Inches(12), Inches(0.4),
        "지역 기반 공동구매 물류 거점 제안",
        size=22, bold=True, color=DARK, font=FONT_TITLE,
    )
    add_rect(s, Inches(0.8), Inches(5.75), Inches(0.6), Inches(0.04), fill=RED)

    # 제안 정보
    add_text(
        s, Inches(0.8), Inches(6.0), Inches(6), Inches(0.3),
        "제안사  ·  제안사명",
        size=12, color=MID, font=FONT_BODY,
    )
    add_text(
        s, Inches(0.8), Inches(6.35), Inches(6), Inches(0.3),
        "제안일  ·  2026. 03. 01",
        size=12, color=MID, font=FONT_BODY,
    )
    add_text(
        s, Inches(0.8), Inches(6.70), Inches(6), Inches(0.3),
        "담당자  ·  010-0000-0000",
        size=12, color=MID, font=FONT_BODY,
    )

    # 우하단 컨셉 태그
    add_rounded(s, SW - Inches(5.2), Inches(6.0), Inches(4.4), Inches(1.1),
                fill=YELLOW_SOFT, radius=0.25)
    add_text(
        s, SW - Inches(5.1), Inches(6.1), Inches(4.2), Inches(0.4),
        "# 매일_방문_이유   # 동네_거점   # 수수료_5%",
        size=11, bold=True, color=RED, font=FONT_TITLE,
    )
    add_text(
        s, SW - Inches(5.1), Inches(6.5), Inches(4.2), Inches(0.5),
        "에브리데이에 가장 잘 어울리는 제안서",
        size=16, bold=True, color=DARK, font=FONT_TITLE,
    )


# ============================================================
# Slide 2 — 오프라인이 흔들리는 시간 (차트 유지)
# ============================================================
def slide_02():
    s = prs.slides.add_slide(BLANK)
    add_left_bar(s)
    add_page_header(s, 2)
    add_headline(
        s, "오프라인이 흔들리는 시간.",
        "온라인 커머스 확장 · 가격 비교 문화 · 매대 중심 차별화 한계. 지역 점포는 신규 방문 이유가 필요하다.",
    )

    # 좌측 요약 불릿
    bullets_y = Inches(2.35)
    bullets = [
        ("01", "온라인 커머스 확장", "오프라인 방문 수 구조적 감소"),
        ("02", "가격 비교 문화", "매장 선택권이 온라인으로 이동"),
        ("03", "매대 중심의 한계", "단순 진열로는 차별화 불가"),
        ("04", "지역 점포의 과제", "신규 방문 이유 설계 필요"),
    ]
    for i, (num, title, desc) in enumerate(bullets):
        yy = bullets_y + Inches(i * 0.85)
        add_text(s, Inches(0.6), yy, Inches(0.6), Inches(0.5),
                 num, size=22, bold=True, color=RED, font=FONT_TITLE)
        add_text(s, Inches(1.3), yy + Inches(0.02), Inches(5), Inches(0.4),
                 title, size=14, bold=True, color=DARK, font=FONT_TITLE)
        add_text(s, Inches(1.3), yy + Inches(0.4), Inches(5), Inches(0.4),
                 desc, size=11, color=MID, font=FONT_BODY)

    # 우측 차트
    chart_data = CategoryChartData()
    chart_data.categories = ['2021', '2022', '2023', '2024', '2025']
    chart_data.add_series('방문 지수', (100, 92, 85, 80, 73))
    cx = Inches(6.8)
    cy = Inches(2.2)
    cw = Inches(6.2)
    ch = Inches(4.3)

    # 차트 배경
    add_rounded(s, cx - Inches(0.2), cy - Inches(0.2), cw + Inches(0.4), ch + Inches(0.4),
                fill=LIGHT, radius=0.05)
    add_text(s, cx, cy - Inches(0.1), Inches(5), Inches(0.4),
             "오프라인 점포 방문 지수 (2021=100)",
             size=12, bold=True, color=DARK, font=FONT_TITLE)

    chart_shape = s.shapes.add_chart(
        XL_CHART_TYPE.LINE_MARKERS, cx, cy + Inches(0.3), cw, ch - Inches(0.3), chart_data
    )
    chart = chart_shape.chart
    chart.has_title = False
    chart.has_legend = False

    # 라인 스타일링
    plot = chart.plots[0]
    for series in plot.series:
        series.format.line.color.rgb = RED
        series.format.line.width = Pt(3)
        try:
            series.marker.format.fill.solid()
            series.marker.format.fill.fore_color.rgb = RED
            series.marker.format.line.color.rgb = RED
            series.marker.size = 8
        except Exception:
            pass
        plot.has_data_labels = True
        dl = plot.data_labels
        dl.font.size = Pt(11)
        dl.font.bold = True
        dl.font.color.rgb = DARK
        dl.position = XL_LABEL_POSITION.ABOVE

    add_text(s, cx, cy + ch + Inches(0.2), Inches(6), Inches(0.3),
             "→ 점포에 다시 방문할 이유를 만드는 구조가 필요합니다.",
             size=12, bold=True, color=RED, font=FONT_TITLE)

    add_footer(s)


# ============================================================
# Slide 3 — 매대 옆, 수령대 하나 (4단계 플로우)
# ============================================================
def slide_03():
    s = prs.slides.add_slide(BLANK)
    add_left_bar(s)
    add_page_header(s, 3)
    add_headline(
        s, "매대 옆, 수령대 하나.",
        "이마트 에브리데이 점포 유휴 공간을 지역 공동구매 수령 거점으로 활용하는 모델.",
    )

    steps = [
        ("01", "소비자 모집", "지역 채널 기반\n온라인 모집", "👥"),
        ("02", "공동구매 진행", "가격·수량 확정\n주문 마감", "🛒"),
        ("03", "점포 납품", "일괄 납품\n별도 택배 없음", "🚚"),
        ("04", "소비자 수령", "QR 확인\n점포에서 픽업", "📍"),
    ]
    start_x = Inches(0.8)
    card_w = Inches(2.6)
    card_h = Inches(3.2)
    gap = Inches(0.4)

    for i, (num, title, desc, icon) in enumerate(steps):
        x = start_x + i * (card_w + gap)
        y = Inches(2.5)
        # 카드 배경
        add_rounded(s, x, y, card_w, card_h, fill=WHITE, radius=0.05)
        add_rect(s, x, y, card_w, Inches(0.08), fill=RED)  # 상단 라인
        # 번호
        add_text(s, x + Inches(0.3), y + Inches(0.3), Inches(1), Inches(0.4),
                 num, size=14, bold=True, color=RED, font=FONT_TITLE)
        # 아이콘
        icon_symbol_box(s, x + Inches(0.9), y + Inches(0.8), Inches(0.9),
                        icon, color=RED, bg=RED_SOFT, border=RED)
        # 제목
        add_text(s, x, y + Inches(2.0), card_w, Inches(0.4),
                 title, size=16, bold=True, color=DARK, font=FONT_TITLE,
                 align=PP_ALIGN.CENTER)
        # 설명
        add_text(s, x + Inches(0.2), y + Inches(2.5), card_w - Inches(0.4), Inches(0.8),
                 desc, size=11, color=MID, font=FONT_BODY, align=PP_ALIGN.CENTER)
        # 카드 테두리
        add_rounded(s, x, y, card_w, card_h, line=BORDER, radius=0.05)

        # 화살표 (마지막 카드 제외)
        if i < len(steps) - 1:
            add_arrow(s, x + card_w + Inches(0.05), y + Inches(1.45),
                      Inches(0.3), Inches(0.3), color=YELLOW)

    add_text(s, Inches(0.8), Inches(6.2), Inches(11.5), Inches(0.4),
             "→ 이마트 에브리데이 점포를 지역 공동구매 상품의 수령 거점으로 활용",
             size=13, bold=True, color=RED, font=FONT_TITLE)
    add_footer(s)


# ============================================================
# Slide 4 — 고정비 0원, 수수료 5%
# ============================================================
def slide_04():
    s = prs.slides.add_slide(BLANK)
    add_left_bar(s)
    add_page_header(s, 4)
    add_headline(
        s, "고정비 0원, 수수료 5%.",
        "공동구매 판매 금액의 5%를 점포 수수료로 지급. 별도 인력·공간 투입 없이 발생하는 순수 수익.",
    )

    # 좌측 핵심 숫자 강조
    add_rounded(s, Inches(0.8), Inches(2.4), Inches(5.5), Inches(4.0),
                fill=RED, radius=0.04)
    add_text(s, Inches(1.0), Inches(2.6), Inches(5), Inches(0.4),
             "POINT", size=11, bold=True, color=YELLOW, font=FONT_TITLE)
    add_text(s, Inches(1.0), Inches(3.0), Inches(5.3), Inches(0.6),
             "판매 금액의",
             size=18, color=WHITE, font=FONT_BODY)
    # 초대형 5%
    tb = s.shapes.add_textbox(Inches(1.0), Inches(3.5), Inches(5.3), Inches(1.6))
    tf = tb.text_frame
    tf.margin_left = Pt(0); tf.margin_right = Pt(0); tf.margin_top = Pt(0); tf.margin_bottom = Pt(0)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    r1 = p.add_run(); r1.text = "5"; r1.font.name = FONT_TITLE; r1.font.size = Pt(140); r1.font.bold = True; r1.font.color.rgb = YELLOW
    r2 = p.add_run(); r2.text = "%"; r2.font.name = FONT_TITLE; r2.font.size = Pt(60); r2.font.bold = True; r2.font.color.rgb = WHITE
    add_text(s, Inches(1.0), Inches(5.3), Inches(5.3), Inches(0.5),
             "를 점포 수수료로 지급", size=18, color=WHITE, font=FONT_BODY)

    # 하단 작은 불릿
    sub = [("고정비", "0 원"), ("인력 투입", "없음"), ("공간", "기존 유휴 공간 활용")]
    for i, (k, v) in enumerate(sub):
        xx = Inches(1.0) + i * Inches(1.75)
        add_text(s, xx, Inches(5.9), Inches(1.7), Inches(0.3),
                 k, size=10, color=YELLOW, font=FONT_BODY)
        add_text(s, xx, Inches(6.2), Inches(1.7), Inches(0.3),
                 v, size=13, bold=True, color=WHITE, font=FONT_TITLE)

    # 우측 예시 테이블 + 미니 차트
    add_rounded(s, Inches(6.7), Inches(2.4), Inches(6.1), Inches(4.0),
                fill=WHITE, line=BORDER, radius=0.04)
    add_text(s, Inches(6.9), Inches(2.6), Inches(5.5), Inches(0.4),
             "EXAMPLE  ·  1회 공구 매출 1,000만 원 기준",
             size=11, bold=True, color=RED, font=FONT_TITLE)

    # 테이블
    rows = [
        ("1회 공구 매출", "1,000만 원", DARK),
        ("점포 수수료 (5%)", "50만 원", RED),
    ]
    for i, (label, val, col) in enumerate(rows):
        ty = Inches(3.2) + i * Inches(0.7)
        add_rect(s, Inches(6.9), ty, Inches(5.7), Inches(0.02), fill=BORDER)
        add_text(s, Inches(6.9), ty + Inches(0.15), Inches(3), Inches(0.5),
                 label, size=14, color=DARK, font=FONT_BODY)
        add_text(s, Inches(9.9), ty + Inches(0.1), Inches(2.7), Inches(0.5),
                 val, size=20, bold=True, color=col, font=FONT_TITLE,
                 align=PP_ALIGN.RIGHT)
    # 하단 라인
    add_rect(s, Inches(6.9), Inches(4.7), Inches(5.7), Inches(0.04), fill=RED)

    # 도넛 차트
    chart_data = CategoryChartData()
    chart_data.categories = ['수수료 (5%)', '나머지 (95%)']
    chart_data.add_series('비율', (5, 95))
    chart_shape = s.shapes.add_chart(
        XL_CHART_TYPE.DOUGHNUT, Inches(7.2), Inches(4.85), Inches(1.6), Inches(1.4), chart_data
    )
    chart = chart_shape.chart
    chart.has_title = False
    chart.has_legend = False
    plot = chart.plots[0]
    # 도넛 조각 색상
    try:
        pts = plot.series[0].points
        pts[0].format.fill.solid(); pts[0].format.fill.fore_color.rgb = RED
        pts[1].format.fill.solid(); pts[1].format.fill.fore_color.rgb = LIGHT
    except Exception:
        pass

    add_text(s, Inches(9.0), Inches(5.0), Inches(3.6), Inches(0.4),
             "점포 수익 50만 원", size=14, bold=True, color=RED, font=FONT_TITLE)
    add_text(s, Inches(9.0), Inches(5.4), Inches(3.6), Inches(0.7),
             "→ 공간도, 인력도 그대로.\n   매달 쌓이는 순수익.",
             size=11, color=MID, font=FONT_BODY, line_spacing=1.4)

    add_footer(s)


# ============================================================
# Slide 5 — 한 번의 장바구니, 두 번의 방문
# ============================================================
def slide_05():
    s = prs.slides.add_slide(BLANK)
    add_left_bar(s)
    add_page_header(s, 5)
    add_headline(
        s, "한 번의 장바구니, 두 번의 방문.",
        "공동구매 수령을 위해 점포 방문이 필수. 방문 중 추가 장바구니가 채워진다.",
    )

    cards = [
        ("🚶", "수령 필수 방문", "공동구매 상품 수령을 위해\n점포 방문이 반드시 필요합니다."),
        ("🛒", "자연스러운 추가 구매", "수령 동선 위 매대에서\n평균 장바구니가 커집니다."),
        ("🔁", "지역 고객 재유입", "공동구매 주기마다\n지역 고객이 반복 방문합니다."),
    ]
    card_w = Inches(3.9)
    card_h = Inches(3.9)
    start_x = Inches(0.8)
    gap = Inches(0.3)

    for i, (icon, title, desc) in enumerate(cards):
        x = start_x + i * (card_w + gap)
        y = Inches(2.4)
        add_rounded(s, x, y, card_w, card_h, fill=LIGHT, radius=0.03)
        add_rect(s, x, y, card_w, Inches(0.08), fill=RED)
        icon_symbol_box(s, x + card_w/2 - Inches(0.55), y + Inches(0.5),
                        Inches(1.1), icon, color=RED, bg=WHITE, border=RED)
        add_text(s, x, y + Inches(2.0), card_w, Inches(0.5),
                 title, size=20, bold=True, color=DARK, font=FONT_TITLE,
                 align=PP_ALIGN.CENTER)
        add_text(s, x + Inches(0.3), y + Inches(2.7), card_w - Inches(0.6), Inches(1.0),
                 desc, size=12, color=MID, font=FONT_BODY, align=PP_ALIGN.CENTER,
                 line_spacing=1.5)

    # 하단 하이라이트 문구
    add_rounded(s, Inches(0.8), Inches(6.6), Inches(11.7), Inches(0.6),
                fill=YELLOW, radius=0.3)
    add_text(s, Inches(0.8), Inches(6.68), Inches(11.7), Inches(0.5),
             "공동구매는 매출 이전에, 방문 유입 전략입니다.",
             size=14, bold=True, color=DARK, font=FONT_TITLE, align=PP_ALIGN.CENTER)

    add_footer(s)


# ============================================================
# Slide 6 — 동네마트만 할 수 있는 일 (비교표)
# ============================================================
def slide_06():
    s = prs.slides.add_slide(BLANK)
    add_left_bar(s)
    add_page_header(s, 6)
    add_headline(
        s, "동네마트만 할 수 있는 일.",
        "유휴 공간 + 지역 고객 + 생활 동선. 에브리데이만의 자산이 그대로 경쟁력이 됩니다.",
    )

    # 비교표 데이터
    headers = ["", "기존 공동구매", "제안 모델"]
    rows_data = [
        ("공간", "임대 공간 필요", "점포 유휴 공간 활용"),
        ("인력", "상주 인력 필요", "점포 수령 구조"),
        ("택배비", "개별 택배비 발생", "일괄 납품"),
        ("운영비", "부담 큼", "수수료 기반 구조"),
    ]

    table_x = Inches(1.0)
    table_y = Inches(2.5)
    col_widths = [Inches(2.5), Inches(4.5), Inches(4.5)]
    row_h = Inches(0.75)

    # 헤더 배경
    add_rect(s, table_x, table_y, sum([w for w in col_widths], Emu(0)),
             Inches(0.8), fill=DARK)
    cur_x = table_x
    for i, h in enumerate(headers):
        color = WHITE if i == 0 else (YELLOW if i == 2 else WHITE)
        add_text(s, cur_x, table_y + Inches(0.22), col_widths[i], Inches(0.4),
                 h, size=14, bold=True, color=color, font=FONT_TITLE,
                 align=PP_ALIGN.CENTER)
        cur_x += col_widths[i]

    # 데이터 행
    for r_idx, row in enumerate(rows_data):
        ry = table_y + Inches(0.8) + r_idx * row_h
        # 배경: 얼룩 없음, 대신 두번째 열만 강조
        bg = LIGHT if r_idx % 2 == 0 else WHITE
        add_rect(s, table_x, ry, sum([w for w in col_widths], Emu(0)), row_h, fill=bg)
        # 제안 모델 열 강조
        add_rect(s, table_x + col_widths[0] + col_widths[1], ry,
                 col_widths[2], row_h, fill=RED_SOFT)

        cur_x = table_x
        for c_idx, cell in enumerate(row):
            if c_idx == 0:
                color = DARK; bold = True; size = 13
            elif c_idx == 1:
                color = MID; bold = False; size = 13
            else:
                color = RED; bold = True; size = 13
            add_text(s, cur_x + Inches(0.2), ry + Inches(0.22),
                     col_widths[c_idx] - Inches(0.4), Inches(0.4),
                     cell, size=size, bold=bold, color=color, font=FONT_BODY,
                     align=PP_ALIGN.CENTER if c_idx > 0 else PP_ALIGN.LEFT)
            cur_x += col_widths[c_idx]

    # 테이블 하단 라인
    total_w = sum([w for w in col_widths], Emu(0))
    total_h = Inches(0.8) + len(rows_data) * row_h
    add_rect(s, table_x, table_y + total_h, total_w, Inches(0.06), fill=RED)

    add_text(s, Inches(1.0), table_y + total_h + Inches(0.25),
             Inches(12), Inches(0.4),
             "→ 점포 유휴 공간을 활용한 비용 최소화 모델.",
             size=13, bold=True, color=RED, font=FONT_TITLE)

    add_footer(s)


# ============================================================
# Slide 7 — 온라인에서 매장까지, 5단계
# ============================================================
def slide_07():
    s = prs.slides.add_slide(BLANK)
    add_left_bar(s)
    add_page_header(s, 7)
    add_headline(
        s, "온라인에서 매장까지, 5단계.",
        "점포 방문에서 시작해 구독·주문·수령까지. 자연스러운 소비자 여정 설계.",
    )

    steps = [
        ("01", "점포 방문", "매장 내 안내 POP"),
        ("02", "QR 스캔", "카운터·매대 QR"),
        ("03", "채널 구독", "지역 공동구매 채널"),
        ("04", "공동구매 신청", "온라인 주문·결제"),
        ("05", "점포 수령", "QR 확인 후 픽업"),
    ]
    card_w = Inches(2.2)
    card_h = Inches(3.0)
    gap = Inches(0.2)
    start_x = Inches(0.8)

    for i, (num, title, desc) in enumerate(steps):
        x = start_x + i * (card_w + gap)
        y = Inches(2.8)
        # 번호 원
        icon_symbol_box(s, x + card_w/2 - Inches(0.5), y, Inches(1.0),
                        num, color=WHITE, bg=RED, border=RED)
        # 카드
        add_rounded(s, x, y + Inches(0.7), card_w, card_h - Inches(0.7),
                    fill=LIGHT, radius=0.04)
        add_text(s, x, y + Inches(1.1), card_w, Inches(0.5),
                 title, size=16, bold=True, color=DARK, font=FONT_TITLE,
                 align=PP_ALIGN.CENTER)
        add_text(s, x + Inches(0.15), y + Inches(1.7), card_w - Inches(0.3), Inches(0.8),
                 desc, size=11, color=MID, font=FONT_BODY, align=PP_ALIGN.CENTER)
        # 화살표
        if i < len(steps) - 1:
            add_arrow(s, x + card_w - Inches(0.05), y + Inches(0.35),
                      Inches(0.25), Inches(0.3), color=YELLOW)

    add_rounded(s, Inches(0.8), Inches(6.4), Inches(11.7), Inches(0.6),
                fill=RED_SOFT, radius=0.3)
    add_text(s, Inches(0.8), Inches(6.48), Inches(11.7), Inches(0.5),
             "오프라인 점포가 온라인 유입의 시작점이 됩니다.",
             size=14, bold=True, color=RED, font=FONT_TITLE, align=PP_ALIGN.CENTER)

    add_footer(s)


# ============================================================
# Slide 8 — 대형마트엔 없는 상품만
# ============================================================
def slide_08():
    s = prs.slides.add_slide(BLANK)
    add_left_bar(s)
    add_page_header(s, 8)
    add_headline(
        s, "대형마트엔 없는 상품만.",
        "가격 경쟁이 아닌 차별화 상품 중심. 이마트 본점 MD와 충돌하지 않는 독자 카테고리.",
    )

    cards = [
        ("💰", "30~40% 할인", "시중가 대비 30~40% 할인된\n공동구매 단가 설계",
         RED),
        ("🏷️", "차별화 상품", "대기업 브랜드가 아닌\n중소·신생 브랜드 중심",
         DARK),
        ("🎁", "선공개 · 지역 한정", "지역 한정 · 선공개 상품으로\n'여기서만' 가치를 만듭니다",
         RED),
        ("📦", "자체 기획 상품", "기획 단계부터 공동구매용으로\n설계된 PB 상품 중심",
         DARK),
    ]
    card_w = Inches(5.8)
    card_h = Inches(2.0)
    gap_x = Inches(0.4)
    gap_y = Inches(0.3)
    start_x = Inches(0.8)
    start_y = Inches(2.4)

    for i, (icon, title, desc, accent) in enumerate(cards):
        col = i % 2
        row = i // 2
        x = start_x + col * (card_w + gap_x)
        y = start_y + row * (card_h + gap_y)
        add_rounded(s, x, y, card_w, card_h, fill=WHITE, line=BORDER, radius=0.03)
        add_rect(s, x, y, Inches(0.1), card_h, fill=accent)
        # 아이콘
        icon_symbol_box(s, x + Inches(0.45), y + Inches(0.45),
                        Inches(1.1), icon, color=accent, bg=RED_SOFT if accent == RED else LIGHT,
                        border=accent)
        # 제목
        add_text(s, x + Inches(1.9), y + Inches(0.35), card_w - Inches(2.2), Inches(0.5),
                 title, size=18, bold=True, color=accent, font=FONT_TITLE)
        # 설명
        add_text(s, x + Inches(1.9), y + Inches(0.95), card_w - Inches(2.2), Inches(1.0),
                 desc, size=12, color=MID, font=FONT_BODY, line_spacing=1.4)

    add_text(s, Inches(0.8), Inches(6.95), Inches(12), Inches(0.3),
             "→ 가격이 아니라 '상품' 자체로 에브리데이에만 있는 이유를 만듭니다.",
             size=13, bold=True, color=RED, font=FONT_TITLE)

    add_footer(s)


# ============================================================
# Slide 9 — 점주는 확인만, 수익은 자동
# ============================================================
def slide_09():
    s = prs.slides.add_slide(BLANK)
    add_left_bar(s)
    add_page_header(s, 9)
    add_headline(
        s, "점주는 확인만, 수익은 자동.",
        "운영 부담을 점주 업무 바깥에 두는 설계. 점포는 상자만 받고, 소비자는 QR로 수령.",
    )

    steps = [
        ("01", "수량 확정", "공동구매 마감\n제안사 측 수량 집계"),
        ("02", "점포 납품", "일괄 배송\n점포 수령대 보관"),
        ("03", "소비자 수령", "QR 확인\n매장에서 픽업"),
        ("04", "5% 정산", "월 단위 정산\n점포 수수료 지급"),
    ]
    card_w = Inches(2.8)
    card_h = Inches(3.5)
    gap = Inches(0.25)
    start_x = Inches(0.9)

    for i, (num, title, desc) in enumerate(steps):
        x = start_x + i * (card_w + gap)
        y = Inches(2.5)
        # 상단 노랑 밴드
        add_rect(s, x, y, card_w, Inches(0.5), fill=YELLOW)
        add_text(s, x, y + Inches(0.08), card_w, Inches(0.35),
                 f"STEP {num}", size=12, bold=True, color=DARK,
                 font=FONT_TITLE, align=PP_ALIGN.CENTER)
        # 카드 본체
        add_rect(s, x, y + Inches(0.5), card_w, card_h - Inches(0.5), fill=WHITE)
        add_rounded(s, x, y, card_w, card_h, line=DARK, radius=0.03)
        # 아이콘 원
        icon_symbol_box(s, x + card_w/2 - Inches(0.6), y + Inches(0.9),
                        Inches(1.2), "✓", color=WHITE, bg=RED, border=RED)
        # 제목
        add_text(s, x, y + Inches(2.3), card_w, Inches(0.5),
                 title, size=18, bold=True, color=DARK, font=FONT_TITLE,
                 align=PP_ALIGN.CENTER)
        add_text(s, x + Inches(0.2), y + Inches(2.85), card_w - Inches(0.4), Inches(0.8),
                 desc, size=11, color=MID, font=FONT_BODY, align=PP_ALIGN.CENTER,
                 line_spacing=1.4)

    add_text(s, Inches(0.8), Inches(6.4), Inches(12), Inches(0.3),
             "→ 점포 운영 부담 최소화 설계. 기존 업무에 끼어들지 않습니다.",
             size=13, bold=True, color=RED, font=FONT_TITLE)

    add_footer(s)


# ============================================================
# Slide 10 — 점주를 귀찮게 하지 않는 설계
# ============================================================
def slide_10():
    s = prs.slides.add_slide(BLANK)
    add_left_bar(s)
    add_page_header(s, 10)
    add_headline(
        s, "점주를 귀찮게 하지 않는 설계.",
        "설명, 상담, 재고관리 없음. QR 한 장이면 끝. 단기 테스트도 바로 가능.",
    )

    items = [
        ("🙅", "설명 인력 불필요", "점주·직원 설명 없음\nQR 코드 안내 중심"),
        ("📱", "QR 안내 구조", "간단한 QR 한 장\n소비자 스스로 접근"),
        ("✅", "단순 수령 확인", "QR 스캔 후 수령 체크\n재고·발주 관리 없음"),
        ("🧪", "단기 테스트 운영", "3~6개월 단기 단위로\n부담 없이 시작 가능"),
    ]
    card_w = Inches(2.8)
    card_h = Inches(3.6)
    gap = Inches(0.25)
    start_x = Inches(0.9)

    for i, (icon, title, desc) in enumerate(items):
        x = start_x + i * (card_w + gap)
        y = Inches(2.5)
        add_rounded(s, x, y, card_w, card_h, fill=RED_SOFT, radius=0.04)
        icon_symbol_box(s, x + card_w/2 - Inches(0.6), y + Inches(0.4),
                        Inches(1.2), icon, color=RED, bg=WHITE, border=RED)
        add_text(s, x, y + Inches(1.9), card_w, Inches(0.5),
                 title, size=17, bold=True, color=DARK, font=FONT_TITLE,
                 align=PP_ALIGN.CENTER)
        add_text(s, x + Inches(0.2), y + Inches(2.45), card_w - Inches(0.4), Inches(1.1),
                 desc, size=11, color=MID, font=FONT_BODY, align=PP_ALIGN.CENTER,
                 line_spacing=1.5)

    add_footer(s)


# ============================================================
# Slide 11 — 3개월, 5개 점포로 증명
# ============================================================
def slide_11():
    s = prs.slides.add_slide(BLANK)
    add_left_bar(s)
    add_page_header(s, 11)
    add_headline(
        s, "3개월, 5개 점포로 증명.",
        "리스크 없는 파일럿 운영 제안. 최소 단위로 데이터를 만들고, 확장은 그 다음입니다.",
    )

    nums = [
        ("3~5", "개 점포", "파일럿 대상 점포 수"),
        ("1~2", "개 상품", "테스트 SKU 수량"),
        ("3", "개월", "운영 기간"),
        ("→", "데이터 검증", "확장 판단 근거"),
    ]
    card_w = Inches(2.8)
    card_h = Inches(3.6)
    gap = Inches(0.25)
    start_x = Inches(0.9)

    for i, (num, unit, desc) in enumerate(nums):
        x = start_x + i * (card_w + gap)
        y = Inches(2.5)
        # 상단 레드 박스
        add_rect(s, x, y, card_w, Inches(2.2), fill=RED)
        # 하단 화이트
        add_rect(s, x, y + Inches(2.2), card_w, Inches(1.4), fill=WHITE)
        add_rounded(s, x, y, card_w, card_h, line=RED, radius=0.04)

        # 초대형 숫자
        tb = s.shapes.add_textbox(x, y + Inches(0.25), card_w, Inches(1.6))
        tf = tb.text_frame
        tf.margin_left = Pt(0); tf.margin_right = Pt(0); tf.margin_top = Pt(0); tf.margin_bottom = Pt(0)
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
        r = p.add_run(); r.text = num; r.font.name = FONT_TITLE
        r.font.size = Pt(72 if num != "→" else 80); r.font.bold = True; r.font.color.rgb = YELLOW
        add_text(s, x, y + Inches(1.75), card_w, Inches(0.4),
                 unit, size=14, bold=True, color=WHITE, font=FONT_TITLE,
                 align=PP_ALIGN.CENTER)

        # 하단 설명
        add_text(s, x + Inches(0.2), y + Inches(2.5), card_w - Inches(0.4), Inches(0.9),
                 desc, size=12, color=DARK, font=FONT_BODY, align=PP_ALIGN.CENTER,
                 line_spacing=1.4)

    add_rounded(s, Inches(0.8), Inches(6.4), Inches(11.7), Inches(0.6),
                fill=YELLOW_SOFT, radius=0.3)
    add_text(s, Inches(0.8), Inches(6.48), Inches(11.7), Inches(0.5),
             "리스크 없는 실험 구조 · 데이터 기반 확장 판단",
             size=14, bold=True, color=RED, font=FONT_TITLE, align=PP_ALIGN.CENTER)

    add_footer(s)


# ============================================================
# Slide 12 — 점포에 남는 것, 다섯 가지
# ============================================================
def slide_12():
    s = prs.slides.add_slide(BLANK)
    add_left_bar(s)
    add_page_header(s, 12)
    add_headline(
        s, "점포에 남는 것, 다섯 가지.",
        "단기 매출을 넘어, 점포가 축적하는 다섯 가지 장기 자산.",
    )

    items = [
        ("🚶", "방문 증가", "지역 고객\n점포 방문 증가"),
        ("💰", "수익 확대", "공동구매 수수료\n점주 수익 확대"),
        ("❤️", "지역 재활성화", "고객 재방문 주기\n지역 고객 재활성화"),
        ("📣", "마케팅 절감", "지역 채널 기반\n마케팅 비용 절감"),
        ("⭐", "경쟁력 강화", "동네마트 포지셔닝\n점포 경쟁력 강화"),
    ]
    card_w = Inches(2.3)
    card_h = Inches(3.6)
    gap = Inches(0.15)
    start_x = Inches(0.8)

    for i, (icon, title, desc) in enumerate(items):
        x = start_x + i * (card_w + gap)
        y = Inches(2.5)
        # 교차로 배경색
        bg = RED if i % 2 == 0 else WHITE
        line_col = RED
        text_col = WHITE if i % 2 == 0 else DARK
        desc_col = YELLOW_SOFT if i % 2 == 0 else MID
        num_col = YELLOW if i % 2 == 0 else RED

        add_rounded(s, x, y, card_w, card_h, fill=bg, line=line_col, radius=0.04)
        add_text(s, x + Inches(0.3), y + Inches(0.3), Inches(1.5), Inches(0.4),
                 f"0{i+1}", size=14, bold=True, color=num_col, font=FONT_TITLE)
        icon_symbol_box(s, x + card_w/2 - Inches(0.55), y + Inches(0.9),
                        Inches(1.1), icon, color=RED, bg=WHITE, border=RED)
        add_text(s, x, y + Inches(2.2), card_w, Inches(0.5),
                 title, size=16, bold=True, color=text_col, font=FONT_TITLE,
                 align=PP_ALIGN.CENTER)
        add_text(s, x + Inches(0.15), y + Inches(2.8), card_w - Inches(0.3), Inches(0.8),
                 desc, size=11, color=desc_col, font=FONT_BODY, align=PP_ALIGN.CENTER,
                 line_spacing=1.4)

    add_footer(s)


# ============================================================
# Slide 13 — 잘못돼도 점포는 손해 없음
# ============================================================
def slide_13():
    s = prs.slides.add_slide(BLANK)
    add_left_bar(s)
    add_page_header(s, 13)
    add_headline(
        s, "잘못돼도 점포는 손해 없음.",
        "고정비·재고·인력 리스크를 모두 제안사 측이 흡수. 점포는 언제든 종료 가능.",
    )

    items = [
        ("🧪", "초기 테스트 운영",
         "제한 점포 · 제한 상품으로\n위험을 최소화",
         "3~5개 점포 / 3개월"),
        ("🛑", "판매 부진 시 종료",
         "판매가 부진하면 즉시 종료\n점포 위약 부담 없음",
         "무조건부 종료권"),
        ("🎯", "사은품 남용 방지",
         "차별화 상품 중심 설계로\n가격 경쟁 · 사은품 의존 제거",
         "상품 경쟁력 기반"),
        ("⚙️", "운영 간소화",
         "설명·재고·발주 없음\nQR 기반 수령 운영",
         "점포 부담 0"),
    ]

    card_w = Inches(5.8)
    card_h = Inches(2.0)
    gap_x = Inches(0.4)
    gap_y = Inches(0.3)
    start_x = Inches(0.8)
    start_y = Inches(2.4)

    for i, (icon, title, desc, tag) in enumerate(items):
        col = i % 2
        row = i // 2
        x = start_x + col * (card_w + gap_x)
        y = start_y + row * (card_h + gap_y)
        add_rounded(s, x, y, card_w, card_h, fill=LIGHT, radius=0.03)
        add_rect(s, x, y, card_w, Inches(0.08), fill=RED)
        icon_symbol_box(s, x + Inches(0.4), y + Inches(0.45),
                        Inches(1.1), icon, color=RED, bg=WHITE, border=RED)
        add_text(s, x + Inches(1.8), y + Inches(0.35), card_w - Inches(2.1), Inches(0.5),
                 title, size=18, bold=True, color=DARK, font=FONT_TITLE)
        add_text(s, x + Inches(1.8), y + Inches(0.9), card_w - Inches(2.1), Inches(0.9),
                 desc, size=11, color=MID, font=FONT_BODY, line_spacing=1.4)
        # 태그
        add_rounded(s, x + Inches(1.8), y + card_h - Inches(0.55),
                    Inches(2.5), Inches(0.35), fill=RED_SOFT, radius=0.4)
        add_text(s, x + Inches(1.8), y + card_h - Inches(0.52),
                 Inches(2.5), Inches(0.3),
                 tag, size=10, bold=True, color=RED, font=FONT_TITLE,
                 align=PP_ALIGN.CENTER)

    add_text(s, Inches(0.8), Inches(6.95), Inches(12), Inches(0.3),
             "→ 점포 리스크 0. 제안사가 먼저 증명하고, 점포는 결과만 공유합니다.",
             size=13, bold=True, color=RED, font=FONT_TITLE)

    add_footer(s)


# ============================================================
# Slide 14 — 에브리데이에 가장 잘 어울리는 제안
# ============================================================
def slide_14():
    s = prs.slides.add_slide(BLANK)
    add_rect(s, 0, 0, SW, SH, fill=WHITE)
    # 상단 밴드
    add_rect(s, 0, 0, SW, Inches(2.0), fill=RED)
    add_rect(s, 0, Inches(2.0), SW, Inches(0.12), fill=YELLOW)

    add_text(s, Inches(0.8), Inches(0.5), Inches(12), Inches(0.4),
             "PROPOSAL CLOSING",
             size=12, bold=True, color=YELLOW, font=FONT_TITLE)
    add_text(s, Inches(0.8), Inches(0.9), Inches(12), Inches(1.2),
             "에브리데이에 가장 잘 어울리는 제안.",
             size=40, bold=True, color=WHITE, font=FONT_TITLE)

    # 4포인트 요약
    points = [
        ("01", "유휴 역량 활용", "점포 유휴 공간으로\n추가 수익 모델 구축"),
        ("02", "방문 유입 중심", "매출 이전에\n방문 유입 구조 설계"),
        ("03", "비용 부담 최소화", "고정비·인력 부담 없이\n수수료 기반 수익"),
        ("04", "파일럿 검증 가능", "3개월 · 5개 점포로\n데이터 기반 확장"),
    ]
    card_w = Inches(2.9)
    card_h = Inches(3.0)
    gap = Inches(0.2)
    start_x = Inches(0.8)

    for i, (num, title, desc) in enumerate(points):
        x = start_x + i * (card_w + gap)
        y = Inches(2.6)
        add_rounded(s, x, y, card_w, card_h, fill=LIGHT, line=BORDER, radius=0.04)
        add_rect(s, x, y, Inches(0.1), card_h, fill=RED)
        add_text(s, x + Inches(0.3), y + Inches(0.4), Inches(2), Inches(0.5),
                 num, size=22, bold=True, color=RED, font=FONT_TITLE)
        add_text(s, x + Inches(0.3), y + Inches(1.0), card_w - Inches(0.5), Inches(0.6),
                 title, size=18, bold=True, color=DARK, font=FONT_TITLE)
        add_rect(s, x + Inches(0.3), y + Inches(1.65), Inches(0.4), Inches(0.04), fill=RED)
        add_text(s, x + Inches(0.3), y + Inches(1.85), card_w - Inches(0.5), Inches(1.0),
                 desc, size=12, color=MID, font=FONT_BODY, line_spacing=1.5)

    # 하단 클로징
    add_rect(s, 0, Inches(6.0), SW, Inches(1.5), fill=DARK)
    add_text(s, Inches(0.8), Inches(6.15), Inches(12), Inches(0.4),
             "이마트 에브리데이 점포에 실질적 이익을 제공하는 구조를 제안드립니다.",
             size=16, color=WHITE, font=FONT_BODY)
    add_text(s, Inches(0.8), Inches(6.55), Inches(12), Inches(0.5),
             "매일, 동네가 모이는 거점.",
             size=28, bold=True, color=YELLOW, font=FONT_TITLE)
    add_text(s, SW - Inches(4.8), Inches(7.0), Inches(4.2), Inches(0.4),
             "제안사명  ·  2026.03.01",
             size=11, color=MID, font=FONT_BODY, align=PP_ALIGN.RIGHT)


# ============================================================
# 빌드 실행
# ============================================================
def main():
    slide_01()
    slide_02()
    slide_03()
    slide_04()
    slide_05()
    slide_06()
    slide_07()
    slide_08()
    slide_09()
    slide_10()
    slide_11()
    slide_12()
    slide_13()
    slide_14()

    out_path = "이마트에브리데이_사업제안서_v2_260301.pptx"
    prs.save(out_path)
    print(f"SAVED: {out_path}")
    print(f"SLIDES: {len(prs.slides)}")


if __name__ == "__main__":
    main()
