# -*- coding: utf-8 -*-
"""
원본 PPT 구조 유지하면서 컬러만 이마트 톤(검정 바탕 + 노랑 + 화이트)으로.

텍스트 색은 '그 텍스트가 올라가 있는 셰이프의 배경 밝기'에 따라 자동 결정:
  - 셰이프에 밝은 fill → 검정 텍스트 (흰 카드 위)
  - 셰이프에 어두운 fill / 노랑 강조 fill → 상황에 맞게 (노랑 위면 검정)
  - 셰이프에 fill 없음 → 슬라이드 배경(검정) 위니까 흰 텍스트

원본 강조(블루)는 노랑 액센트로 후처리.
"""
import sys
import shutil
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt, Emu
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN

sys.stdout.reconfigure(encoding='utf-8')

SRC = "이마트에브리데이 사업제안서_260301.pptx"
OUT = "이마트에브리데이_사업제안서_v2_260301.pptx"

# ============================================================
# 이마트 컬러
# ============================================================
YELLOW = RGBColor(0xFF, 0xD2, 0x00)
YELLOW_SOFT = RGBColor(0xFF, 0xEA, 0x80)
BLACK = RGBColor(0x15, 0x15, 0x15)
GRAY_MID = RGBColor(0x55, 0x55, 0x55)
LIGHT_GRAY = RGBColor(0xBB, 0xBB, 0xBB)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

# ============================================================
# 원본 FILL 매핑
# ============================================================
FILL_MAP = {
    "002B5B": YELLOW,       # 원본 네이비 강조 블록 → 노랑
    "1F6BBA": YELLOW,       # 원본 블루 강조 → 노랑
    "74A9D8": YELLOW_SOFT,  # 연블루 → 연노랑
    "F0F4F8": WHITE,        # 연한 카드 배경 → 흰색
    "FFFFFF": WHITE,
}

LINE_MAP = {
    "002B5B": YELLOW,
    "1F6BBA": YELLOW,
    "74A9D8": LIGHT_GRAY,
    "DDDDDD": LIGHT_GRAY,
    "E5E5E5": LIGHT_GRAY,
}

# 원본 텍스트 컬러 중 "강조"로 간주되어 노랑으로 바꿀 것들
ACCENT_TEXT_COLORS = {"1F6BBA"}


# ============================================================
# Fill 리컬러
# ============================================================
def recolor_fill(shape):
    try:
        if shape.fill.type == 1:
            hex_col = str(shape.fill.fore_color.rgb).upper()
            if hex_col in FILL_MAP:
                shape.fill.fore_color.rgb = FILL_MAP[hex_col]
    except Exception:
        pass


def recolor_line(shape):
    try:
        if shape.line.fill.type == 1:
            hex_col = str(shape.line.color.rgb).upper()
            if hex_col in LINE_MAP:
                shape.line.color.rgb = LINE_MAP[hex_col]
    except Exception:
        pass


# ============================================================
# 셰이프의 '배경 밝기' 판별
# ============================================================
def luminance(rgb_hex):
    r = int(rgb_hex[0:2], 16)
    g = int(rgb_hex[2:4], 16)
    b = int(rgb_hex[4:6], 16)
    return 0.299 * r + 0.587 * g + 0.114 * b


def shape_own_fill_light(shape):
    """이 셰이프 자신에 솔리드 fill이 있고 밝은 색이면 True."""
    try:
        if shape.fill.type == 1:
            hex_col = str(shape.fill.fore_color.rgb).upper()
            return luminance(hex_col) > 170
    except Exception:
        pass
    return None


def _bbox(shape):
    try:
        x, y = shape.left, shape.top
        w, h = shape.width, shape.height
        if x is None or y is None or w is None or h is None:
            return None
        return (x, y, x + w, y + h, max(1, w * h))
    except Exception:
        return None


def text_sits_on_light_block(shape, z_below_shapes):
    """텍스트 셰이프 위치가 z-order상 아래에 깔린 밝은 fill 셰이프와 50% 이상 겹치는가."""
    tb = _bbox(shape)
    if tb is None:
        return False
    tx1, ty1, tx2, ty2, tarea = tb
    for other in z_below_shapes:
        if other is shape:
            continue
        try:
            if other.fill.type != 1:
                continue
            hex_col = str(other.fill.fore_color.rgb).upper()
            if luminance(hex_col) <= 170:
                continue
        except Exception:
            continue
        ob = _bbox(other)
        if ob is None:
            continue
        ox1, oy1, ox2, oy2, _ = ob
        ix = max(0, min(tx2, ox2) - max(tx1, ox1))
        iy = max(0, min(ty2, oy2) - max(ty1, oy1))
        if ix * iy / tarea > 0.5:
            return True
    return False


def decide_text_color(shape, original_hex, z_below):
    bg_light = shape_own_fill_light(shape)
    if bg_light is None:
        # 자기 fill 없음 → z-order상 아래 셰이프 체크
        if text_sits_on_light_block(shape, z_below):
            bg_light = True
        else:
            bg_light = False

    is_accent = original_hex and original_hex.upper() in ACCENT_TEXT_COLORS
    if is_accent:
        return BLACK if bg_light else YELLOW
    return BLACK if bg_light else WHITE


def recolor_text_in_shape(shape, z_below):
    if not shape.has_text_frame:
        return
    for para in shape.text_frame.paragraphs:
        for run in para.runs:
            orig = None
            try:
                if run.font.color and run.font.color.type is not None:
                    orig = str(run.font.color.rgb).upper()
            except Exception:
                orig = None
            color = decide_text_color(shape, orig, z_below)
            try:
                run.font.color.rgb = color
            except Exception:
                pass


# ============================================================
# 테이블 셀 처리
# ============================================================
def recolor_table(shape):
    if not shape.has_table:
        return
    tbl = shape.table
    for row in tbl.rows:
        for cell in row.cells:
            # 셀 fill 리컬러
            try:
                if cell.fill.type == 1:
                    hex_col = str(cell.fill.fore_color.rgb).upper()
                    if hex_col in FILL_MAP:
                        cell.fill.fore_color.rgb = FILL_MAP[hex_col]
            except Exception:
                pass

            # 셀 배경 밝기로 텍스트 색 결정
            cell_bg_light = None
            try:
                if cell.fill.type == 1:
                    hex_col = str(cell.fill.fore_color.rgb).upper()
                    cell_bg_light = luminance(hex_col) > 170
            except Exception:
                pass

            text_color = BLACK if cell_bg_light else WHITE
            for para in cell.text_frame.paragraphs:
                for run in para.runs:
                    orig = None
                    try:
                        if run.font.color and run.font.color.type is not None:
                            orig = str(run.font.color.rgb).upper()
                    except Exception:
                        pass
                    if orig and orig in ACCENT_TEXT_COLORS and not cell_bg_light:
                        run.font.color.rgb = YELLOW
                    else:
                        try:
                            run.font.color.rgb = text_color
                        except Exception:
                            pass


# ============================================================
# 셰이프 트리 순회 (두 패스: 1) fill/line 전부 리컬러 → 2) 텍스트 리컬러)
# ============================================================
def _flatten(shapes):
    out = []
    for shape in shapes:
        if shape.shape_type == 6:
            try:
                out.extend(_flatten(shape.shapes))
                continue
            except Exception:
                pass
        out.append(shape)
    return out


def process_shapes(shapes):
    flat = _flatten(shapes)
    # Pass 1: fill/line 리컬러
    for shape in flat:
        recolor_fill(shape)
        recolor_line(shape)
        recolor_table(shape)
    # Pass 2: 텍스트 리컬러 (자기 아래 z-order 셰이프 참조)
    for i, shape in enumerate(flat):
        z_below = flat[:i]
        recolor_text_in_shape(shape, z_below)


# ============================================================
# 차트 리컬러
# ============================================================
def recolor_chart(shape):
    if not shape.has_chart:
        return
    chart = shape.chart
    try:
        for plot in chart.plots:
            for series in plot.series:
                try:
                    series.format.line.color.rgb = YELLOW
                    series.format.line.width = Pt(3.5)
                except Exception:
                    pass
                try:
                    series.marker.format.fill.solid()
                    series.marker.format.fill.fore_color.rgb = YELLOW
                    series.marker.format.line.color.rgb = WHITE
                    series.marker.size = 9
                except Exception:
                    pass
                try:
                    pts = list(series.points)
                    palette = [YELLOW, WHITE, YELLOW_SOFT, LIGHT_GRAY]
                    for i, pt in enumerate(pts):
                        pt.format.fill.solid()
                        pt.format.fill.fore_color.rgb = palette[i % len(palette)]
                        try:
                            pt.format.line.color.rgb = BLACK
                        except Exception:
                            pass
                except Exception:
                    pass
                try:
                    plot.has_data_labels = True
                    dl = plot.data_labels
                    dl.font.bold = True
                    dl.font.color.rgb = WHITE
                    dl.font.size = Pt(11)
                except Exception:
                    pass
        try:
            for ax in (chart.category_axis, chart.value_axis):
                ax.tick_labels.font.color.rgb = WHITE
                ax.tick_labels.font.size = Pt(10)
        except Exception:
            pass
    except Exception as e:
        print(f"  chart recolor warn: {e}")


# ============================================================
# 배경 검정화
# ============================================================
def set_black_background(slide):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = BLACK


# ============================================================
# 이마트 프레임 (좌측 노랑바 + 상단 띠 + 페이지번호)
# ============================================================
def add_rect(slide, x, y, w, h, fill):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    shp.shadow.inherit = False
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    shp.line.fill.background()
    return shp


def add_text(slide, x, y, w, h, text, *, size=10, bold=True, color=YELLOW,
             align=PP_ALIGN.LEFT):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.margin_left = Pt(0); tf.margin_right = Pt(0)
    tf.margin_top = Pt(0); tf.margin_bottom = Pt(0)
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = color
    return tb


def add_emart_frame(slide, idx, total, sw, sh):
    add_rect(slide, 0, 0, Inches(0.08), sh, YELLOW)     # 좌측 세로바
    add_rect(slide, 0, 0, sw, Inches(0.08), YELLOW)     # 상단 띠
    add_text(slide, Inches(0.2), Inches(0.2), Inches(5), Inches(0.3),
             "EMART EVERYDAY  ·  사업제안", size=9, color=YELLOW)
    add_text(slide, sw - Inches(1.5), Inches(0.2), Inches(1.2), Inches(0.3),
             f"{idx:02d} / {total:02d}", size=9, color=YELLOW, align=PP_ALIGN.RIGHT)


def style_cover(slide, sw, sh):
    add_rect(slide, 0, 0, sw, Inches(0.25), YELLOW)
    add_rect(slide, 0, sh - Inches(0.15), sw, Inches(0.15), YELLOW)


def style_closing(slide, sw, sh):
    add_rect(slide, 0, 0, sw, Inches(0.15), YELLOW)
    add_rect(slide, 0, sh - Inches(0.25), sw, Inches(0.25), YELLOW)


# ============================================================
# 메인
# ============================================================
def main():
    shutil.copy(SRC, OUT)
    prs = Presentation(OUT)
    sw = prs.slide_width
    sh = prs.slide_height
    total = len(prs.slides)

    for idx, slide in enumerate(prs.slides, start=1):
        set_black_background(slide)
        process_shapes(slide.shapes)
        for shape in list(slide.shapes):
            recolor_chart(shape)
        add_emart_frame(slide, idx, total, sw, sh)
        if idx == 1:
            style_cover(slide, sw, sh)
        elif idx == total:
            style_closing(slide, sw, sh)

    prs.save(OUT)
    print(f"SAVED: {OUT}")
    print(f"SLIDES: {total}")


if __name__ == "__main__":
    main()
