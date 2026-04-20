# -*- coding: utf-8 -*-
"""
A안 보존형 리컬러: 원본 흰 배경/레이아웃 유지, 파란 계열 강조만 노랑+검정테두리.

입력:  이마트에브리데이 사업제안서_260301.pptx
출력:  이마트에브리데이_사업제안서_v3_260301.pptx

규칙 요약 (design doc: docs/superpowers/specs/2026-04-20-emart-ppt-blue-to-yellow-accent-design.md):
- Fill 파랑(#002B5B, #1F6BBA, #74A9D8) → 노랑(#FFD200) + 검정 테두리(#151515, 1.25pt)
- Line 파랑 → 검정
- 텍스트 원본 유지. 단, 파란 블록 안에 있던 흰 텍스트만 검정으로 교체(가독성)
- 배경/프레임/페이지번호 변경 없음
"""
import os
import sys
import shutil
from io import BytesIO
from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt
from pptx.enum.text import MSO_AUTO_SIZE
from PIL import Image, ImageFilter

sys.stdout.reconfigure(encoding='utf-8')

SRC = "이마트에브리데이 사업제안서_260301.pptx"
OUT = "이마트에브리데이_사업제안서_v3_260301.pptx"

YELLOW = RGBColor(0xFF, 0xD2, 0x00)
BLACK = RGBColor(0x15, 0x15, 0x15)

BLUE_FILLS = {"002B5B", "1F6BBA", "74A9D8"}
BLUE_LINES = {"002B5B", "1F6BBA", "74A9D8"}
BORDER_WIDTH = Pt(1.25)


def is_blueish(hx):
    """파랑 우위 색 판정 (해석 못 한 파란 계열 텍스트까지 잡기 위함)."""
    if not hx or len(hx) != 6:
        return False
    try:
        r = int(hx[0:2], 16); g = int(hx[2:4], 16); b = int(hx[4:6], 16)
    except Exception:
        return False
    if b < 80:
        return False
    return b > r + 20 and b > g + 10


def _hex(rgb):
    return str(rgb).upper() if rgb is not None else None


def was_blue_fill(shape_or_cell):
    """변경 전 fill이 파란 계열이었는지 판정 (호출 시점의 현재 상태 기준)."""
    try:
        if shape_or_cell.fill.type == 1:
            return _hex(shape_or_cell.fill.fore_color.rgb) in BLUE_FILLS
    except Exception:
        pass
    return False


def recolor_fill_and_add_border(shape):
    """파란 fill → 노랑 + 검정 테두리. 반환: 변경됐으면 True."""
    try:
        if shape.fill.type != 1:
            return False
        hex_col = _hex(shape.fill.fore_color.rgb)
        if hex_col not in BLUE_FILLS:
            return False
        shape.fill.fore_color.rgb = YELLOW
        try:
            shape.line.color.rgb = BLACK
            shape.line.width = BORDER_WIDTH
        except Exception:
            pass
        return True
    except Exception:
        return False


def recolor_line_only(shape):
    """fill과 무관하게 도형 외곽선이 파란 계열이면 검정으로 교체."""
    try:
        if shape.line.fill.type == 1:
            hex_col = _hex(shape.line.color.rgb)
            if hex_col in BLUE_LINES:
                shape.line.color.rgb = BLACK
    except Exception:
        pass


def recolor_text_on_yellow_block(shape):
    """파란 블록(이제 노랑)이었던 셰이프 안의 흰 텍스트만 검정으로."""
    if not shape.has_text_frame:
        return
    for para in shape.text_frame.paragraphs:
        for run in para.runs:
            try:
                if run.font.color and run.font.color.type is not None:
                    if _hex(run.font.color.rgb) == "FFFFFF":
                        run.font.color.rgb = BLACK
            except Exception:
                pass


def recolor_blue_text(shape):
    """파란 계열 텍스트 → 검정 (셰이프 전체 대상, 블록 바깥 포함)."""
    if not shape.has_text_frame:
        return
    for para in shape.text_frame.paragraphs:
        for run in para.runs:
            try:
                if run.font.color and run.font.color.type is not None:
                    if is_blueish(_hex(run.font.color.rgb)):
                        run.font.color.rgb = BLACK
            except Exception:
                pass


def recolor_table(shape):
    if not shape.has_table:
        return
    for row in shape.table.rows:
        for cell in row.cells:
            # 셀 fill 파랑 → 노랑 + 검정 테두리 흉내 (python-pptx는 셀 border 직접 지원 제한적
            # → fill 위주로 처리)
            changed = False
            try:
                if cell.fill.type == 1 and _hex(cell.fill.fore_color.rgb) in BLUE_FILLS:
                    cell.fill.fore_color.rgb = YELLOW
                    changed = True
            except Exception:
                pass
            # 셀 안 텍스트 정리:
            #   - 이 셀이 파랑이었다면 흰 텍스트 → 검정 (가독성)
            #   - 파란 계열 텍스트 → 검정 (셀 배경과 무관)
            for para in cell.text_frame.paragraphs:
                for run in para.runs:
                    try:
                        if run.font.color and run.font.color.type is not None:
                            hx = _hex(run.font.color.rgb)
                            if changed and hx == "FFFFFF":
                                run.font.color.rgb = BLACK
                            elif is_blueish(hx):
                                run.font.color.rgb = BLACK
                    except Exception:
                        pass


def recolor_chart(shape):
    if not shape.has_chart:
        return
    try:
        for plot in shape.chart.plots:
            for series in plot.series:
                # 라인
                try:
                    if _hex(series.format.line.color.rgb) in BLUE_FILLS:
                        series.format.line.color.rgb = YELLOW
                except Exception:
                    pass
                # 마커
                try:
                    if _hex(series.marker.format.fill.fore_color.rgb) in BLUE_FILLS:
                        series.marker.format.fill.solid()
                        series.marker.format.fill.fore_color.rgb = YELLOW
                except Exception:
                    pass
                # 포인트(막대 등)
                try:
                    for pt in series.points:
                        try:
                            if _hex(pt.format.fill.fore_color.rgb) in BLUE_FILLS:
                                pt.format.fill.solid()
                                pt.format.fill.fore_color.rgb = YELLOW
                        except Exception:
                            continue
                except Exception:
                    pass
    except Exception as e:
        print(f"  chart warn: {e}")


def set_text_outline(run, rgb_hex="151515", width_pt=0.75):
    """런 텍스트에 검정 아웃라인(stroke) 추가 — 노란 위 흰 글씨 가독성용."""
    rPr = run._r.get_or_add_rPr()
    for old in rPr.findall(qn("a:ln")):
        rPr.remove(old)
    ln = etree.SubElement(rPr, qn("a:ln"))
    ln.set("w", str(int(round(width_pt * 12700))))  # 1pt = 12700 EMU
    solid = etree.SubElement(ln, qn("a:solidFill"))
    clr = etree.SubElement(solid, qn("a:srgbClr"))
    clr.set("val", rgb_hex)
    # CT_TextCharacterProperties 스키마상 a:ln 은 rPr 의 첫 자식
    rPr.remove(ln)
    rPr.insert(0, ln)


def collect_yellow_bboxes(slide):
    """슬라이드 내 모든 노란 fill 셰이프의 (shape, bbox) 리스트."""
    out = []
    for s in _flatten(slide.shapes):
        try:
            if s.fill.type == 1 and _hex(s.fill.fore_color.rgb) == "FFD200":
                if s.left is None or s.top is None or s.width is None or s.height is None:
                    continue
                out.append((s, s.left, s.top, s.left + s.width, s.top + s.height))
        except Exception:
            continue
    return out


def _shape_bbox(shape):
    try:
        x, y, w, h = shape.left, shape.top, shape.width, shape.height
        if None in (x, y, w, h):
            return None
        return x, y, x + w, y + h, max(1, w * h)
    except Exception:
        return None


def strip_text_outline(run):
    """기존에 들어갔던 a:ln 아웃라인 제거."""
    rPr = run._r.find(qn("a:rPr"))
    if rPr is None:
        return
    for old in rPr.findall(qn("a:ln")):
        rPr.remove(old)


def blacken_white_text_on_yellow(slide, yellow_bboxes):
    """
    노랑 bbox 와 >=50% 겹치는 셰이프(자기 fill 노랑 제외) 안의 흰 텍스트 → 검정.
    과거에 들어간 아웃라인이 있으면 제거.
    """
    if not yellow_bboxes:
        return
    for shape in _flatten(slide.shapes):
        if not shape.has_text_frame:
            continue
        try:
            if shape.fill.type == 1 and _hex(shape.fill.fore_color.rgb) == "FFD200":
                continue
        except Exception:
            pass
        bb = _shape_bbox(shape)
        if bb is None:
            continue
        sx1, sy1, sx2, sy2, sarea = bb
        hit = False
        for _, yx1, yy1, yx2, yy2 in yellow_bboxes:
            ix = max(0, min(sx2, yx2) - max(sx1, yx1))
            iy = max(0, min(sy2, yy2) - max(sy1, yy1))
            if ix * iy / sarea >= 0.5:
                hit = True
                break
        if not hit:
            continue
        for para in shape.text_frame.paragraphs:
            for run in para.runs:
                strip_text_outline(run)
                try:
                    if run.font.color and run.font.color.type is not None:
                        if _hex(run.font.color.rgb) == "FFFFFF":
                            run.font.color.rgb = BLACK
                except Exception:
                    pass


def find_container(pic, yellow_bboxes):
    """그림의 중심을 포함하는 가장 작은 노란 셰이프 반환 (혹은 None)."""
    bb = _shape_bbox(pic)
    if bb is None:
        return None
    sx1, sy1, sx2, sy2, _ = bb
    cx = (sx1 + sx2) / 2
    cy = (sy1 + sy2) / 2
    best = None
    best_area = None
    for ys, yx1, yy1, yx2, yy2 in yellow_bboxes:
        if ys is pic:
            continue
        if yx1 <= cx <= yx2 and yy1 <= cy <= yy2:
            area = (yx2 - yx1) * (yy2 - yy1)
            if best is None or area < best_area:
                best = ys
                best_area = area
    return best


def fit_icon_in_container(pic, container):
    """
    아이콘을 컨테이너에 맞춰 리사이즈 + 종횡비 보정 + 중앙 정렬.
    - 목표 occupy = min(현재 occupy, 0.60) — 이미 여유 있으면 유지, 꽉 차면 줄임
    - 원본 PNG 종횡비를 유지하면서 컨테이너의 target_occupy 박스 안에 맞춤
    - 거의 변화 없으면(< 2% 오차) 건드리지 않음
    """
    try:
        im = Image.open(BytesIO(pic.image.blob))
        iw, ih = im.size
    except Exception:
        return False
    if iw <= 0 or ih <= 0:
        return False
    native_ar = iw / ih

    cw, ch = container.width, container.height
    ccx = container.left + cw / 2
    ccy = container.top + ch / 2

    cur_w, cur_h = pic.width, pic.height
    cur_occupy = max(cur_w / cw, cur_h / ch)
    target_occupy = min(cur_occupy, 0.60)

    # 컨테이너 정사각 기준 target box
    target_max_w = cw * target_occupy
    target_max_h = ch * target_occupy
    # 원본 종횡비 맞게 축약
    box_ar = target_max_w / target_max_h
    if native_ar > box_ar:
        new_w = target_max_w
        new_h = target_max_w / native_ar
    else:
        new_h = target_max_h
        new_w = target_max_h * native_ar

    # 거의 같으면 skip
    if (abs(new_w - cur_w) < cur_w * 0.02) and (abs(new_h - cur_h) < cur_h * 0.02):
        return False

    pic.width = int(round(new_w))
    pic.height = int(round(new_h))
    pic.left = int(round(ccx - new_w / 2))
    pic.top = int(round(ccy - new_h / 2))
    return True


def reposition_icons(slide, yellow_bboxes):
    """노란 컨테이너를 가진 PICTURE 를 종횡비 복원 + 과밀시 축소."""
    for shape in _flatten(slide.shapes):
        if shape.shape_type != 13:  # PICTURE
            continue
        container = find_container(shape, yellow_bboxes)
        if container is None:
            continue
        fit_icon_in_container(shape, container)


def recolor_png_blue_to_black(blob):
    """
    단색 파란 아이콘(PNG)을 검정으로. 반환: (new_blob, did_change).
    사진 등 복합 이미지는 가드로 스킵:
      - 파란 픽셀 비율이 5% 이상
      - 파랑/투명이 아닌 다른 색 픽셀이 5% 미만
    안티에일리어싱된 반투명 파랑 가장자리까지 포함해 RGB만 검정으로 치환,
    알파는 그대로 유지.
    """
    try:
        im = Image.open(BytesIO(blob))
    except Exception:
        return blob, False
    if im.mode != "RGBA":
        try:
            im = im.convert("RGBA")
        except Exception:
            return blob, False

    pixels = list(im.getdata())
    total = len(pixels)
    if total == 0:
        return blob, False

    blue = other = 0
    for r, g, b, a in pixels:
        if a < 10:
            continue
        if b >= 80 and b > r + 20 and b > g + 10:
            blue += 1
        else:
            other += 1

    if blue * 100 < 5 * total:  # 파란 비율 5% 미만이면 아이콘 아님
        return blob, False
    if other * 100 >= 5 * total:  # 사진/복합 이미지
        return blob, False

    new_pixels = []
    for r, g, b, a in pixels:
        if a >= 10 and b >= 80 and b > r + 20 and b > g + 10:
            new_pixels.append((0x15, 0x15, 0x15, a))
        else:
            new_pixels.append((r, g, b, a))
    im.putdata(new_pixels)
    buf = BytesIO()
    im.save(buf, format="PNG")
    return buf.getvalue(), True


def recolor_all_picture_parts(prs):
    """프레젠테이션 내부 모든 image part를 순회하며 파란 아이콘을 검정화."""
    changed = 0
    total = 0
    for part in prs.part.package.iter_parts():
        ct = getattr(part, "content_type", "") or ""
        if not ct.startswith("image/png"):
            continue
        total += 1
        new_blob, did = recolor_png_blue_to_black(part.blob)
        if did:
            part._blob = new_blob
            changed += 1
    print(f"  icons: {changed}/{total} PNG parts recolored")


def _fit_crop(img, w_px, h_px):
    sw, sh = img.size
    tgt_ar = w_px / h_px
    src_ar = sw / sh
    if src_ar >= tgt_ar:
        scale = h_px / sh
        new_w = int(round(sw * scale))
        img = img.resize((new_w, h_px), Image.LANCZOS)
        left = (new_w - w_px) // 2
        return img.crop((left, 0, left + w_px, h_px))
    else:
        scale = w_px / sw
        new_h = int(round(sh * scale))
        img = img.resize((w_px, new_h), Image.LANCZOS)
        top = (new_h - h_px) // 2
        return img.crop((0, top, w_px, top + h_px))


def _duotone(img, dark, light):
    gray = img.convert("L")
    lut_r = [int(dark[0] + (light[0] - dark[0]) * i / 255) for i in range(256)]
    lut_g = [int(dark[1] + (light[1] - dark[1]) * i / 255) for i in range(256)]
    lut_b = [int(dark[2] + (light[2] - dark[2]) * i / 255) for i in range(256)]
    r = gray.point(lut_r)
    g = gray.point(lut_g)
    b = gray.point(lut_b)
    return Image.merge("RGB", (r, g, b))


def make_cover_background(src_path, out_path, w_px, h_px):
    """
    표지 우측 배경: 이마트 간판 사진 → 노랑/검정 듀오톤 + 약한 블러 + 좌측 엣지 노랑 페이드.
    알아볼 수 있는 수준의 약한 블러 (<0.3% of height).
    """
    img = Image.open(src_path).convert("RGB")
    img = _fit_crop(img, w_px, h_px)

    # 듀오톤 (어두움=검정, 밝음=노랑) — 원본 파랑간판의 파랑은 어두운 톤이라 검정으로,
    # 노랑 "everyday" 는 그대로 노랑 유지 (브랜드 컬러 일관)
    img = _duotone(img, (0x10, 0x10, 0x10), (0xFF, 0xD2, 0x00))

    # 약한 블러 (로고 글자 형태는 인식 가능)
    img = img.filter(ImageFilter.GaussianBlur(radius=max(2, int(h_px * 0.0025))))

    # 좌측 20% 영역 → 노랑으로 자연스럽게 페이드 (좌측 노랑 블록과 연결감)
    fade_w = int(w_px * 0.20)
    fade_row = Image.new("L", (w_px, 1), 255)
    for x in range(fade_w):
        fade_row.putpixel((x, 0), int(255 * (x / max(fade_w - 1, 1))))
    fade = fade_row.resize((w_px, h_px), Image.BILINEAR)
    yellow_bg = Image.new("RGB", (w_px, h_px), (0xFF, 0xD2, 0x00))
    img = Image.composite(img, yellow_bg, fade)

    img.save(out_path, "PNG")


def move_shape_to_back(shape):
    """해당 셰이프 요소를 spTree 맨 앞으로 옮겨서 z-order 최하단으로."""
    sp = shape._element
    spTree = sp.getparent()
    spTree.remove(sp)
    # spTree 의 첫 두 자식(nvGrpSpPr, grpSpPr) 뒤, 즉 인덱스 2 위치에 삽입
    spTree.insert(2, sp)


def _add_rect(slide, x, y, w, h, color):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    shp.fill.solid()
    shp.fill.fore_color.rgb = color
    shp.line.fill.background()
    shp.shadow.inherit = False
    return shp


def _add_textbox(slide, x, y, w, h, text, *, size_pt, bold=False,
                 color=None, align=None, font_name=None):
    from pptx.enum.text import PP_ALIGN
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.word_wrap = True
    tf.margin_left = Emu(0)
    tf.margin_right = Emu(0)
    tf.margin_top = Emu(0)
    tf.margin_bottom = Emu(0)
    p = tf.paragraphs[0]
    if align is not None:
        p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.size = Pt(size_pt)
    r.font.bold = bold
    if color is not None:
        r.font.color.rgb = color
    if font_name:
        r.font.name = font_name
    return tb


def style_cover_slide(slide, prs, bg_path):
    """
    에디토리얼/브루탈리즘 스타일 표지:
    - 좌측 절반 노랑 솔리드 / 우측 절반 간판 듀오톤(브랜드 컬러 유지) + 좌측 페이드
    - 상·하단 얇은 검정 바 (strong grid)
    - 좌상단 작은 메타 레이블, 우상단 대형 "NO.01"
    - 중앙 좌측에 대형 타이틀 한 줄 + 서브타이틀
    - 좌하단 얇은 구분선 + 3줄 메타 정보
    """
    from pptx.enum.text import PP_ALIGN

    sw = prs.slide_width
    sh = prs.slide_height
    half = sw // 2

    def _in(x):  # inches → EMU
        return Emu(int(round(x * 914400)))

    # 1) 좌측 절반 노랑 / 우측 절반 사진 (둘 다 맨 뒤)
    left_rect = _add_rect(slide, 0, 0, half, sh, YELLOW)
    move_shape_to_back(left_rect)
    pic = slide.shapes.add_picture(bg_path, half, 0, sw - half, sh)
    move_shape_to_back(pic)

    # 2) 상단 두꺼운 검정 바 + 하단 얇은 검정 바
    top_bar_h = _in(0.38)
    _add_rect(slide, 0, 0, sw, top_bar_h, BLACK)
    _add_rect(slide, 0, sh - _in(0.08), sw, _in(0.08), BLACK)

    # 3) 상단 바 좌측: 메타 레이블 (노랑)
    _add_textbox(
        slide,
        _in(0.40), _in(0.08), _in(7.0), _in(0.25),
        "EMART EVERYDAY   ·   PROPOSAL   ·   2026.03",
        size_pt=10, bold=True, color=YELLOW, font_name="Consolas",
    )

    # 4) 상단 바 우측: NO.01 (노랑)
    _add_textbox(
        slide,
        _in(7.6), _in(0.06), _in(2.0), _in(0.28),
        "NO.01 / 01", size_pt=12, bold=True, color=YELLOW,
        align=PP_ALIGN.RIGHT, font_name="Consolas",
    )

    # 5) 좌측 중앙 가로 포인트 라인
    _add_rect(slide, _in(0.42), _in(1.75), _in(0.6), _in(0.04), BLACK)

    # 6) 기존 Text 1~4 재스타일링
    for shape in list(slide.shapes):
        if not shape.has_text_frame:
            continue
        name = shape.name
        if name == "Text 1":
            shape.left = _in(0.40)
            shape.top = _in(2.00)
            shape.width = _in(4.50)
            shape.height = _in(2.40)
            tf = shape.text_frame
            tf.auto_size = MSO_AUTO_SIZE.NONE
            tf.word_wrap = True
            tf.margin_left = Emu(0)
            tf.margin_right = Emu(0)
            tf.margin_top = Emu(0)
            tf.margin_bottom = Emu(0)
            for i, para in enumerate(tf.paragraphs):
                for run in para.runs:
                    strip_text_outline(run)
                    run.font.color.rgb = BLACK
                    if i == 0:
                        run.font.bold = True
                        run.font.size = Pt(26)
                    else:
                        run.font.bold = False
                        run.font.size = Pt(12)
        elif name in ("Text 2", "Text 3", "Text 4"):
            idx = {"Text 2": 0, "Text 3": 1, "Text 4": 2}[name]
            shape.left = _in(0.40)
            shape.top = _in(4.75 + 0.23 * idx)
            shape.width = _in(5.00)
            shape.height = _in(0.22)
            tf = shape.text_frame
            tf.margin_left = Emu(0)
            tf.margin_right = Emu(0)
            tf.margin_top = Emu(0)
            tf.margin_bottom = Emu(0)
            for para in tf.paragraphs:
                for run in para.runs:
                    strip_text_outline(run)
                    run.font.color.rgb = BLACK
                    run.font.size = Pt(10)

    # 7) 하단 메타 위 얇은 구분선
    _add_rect(slide, _in(0.42), _in(4.60), _in(0.8), _in(0.03), BLACK)


def _flatten(shapes):
    out = []
    for shape in shapes:
        if shape.shape_type == 6:  # GROUP
            try:
                out.extend(_flatten(shape.shapes))
                continue
            except Exception:
                pass
        out.append(shape)
    return out


def process_slide(slide):
    flat = _flatten(slide.shapes)
    for shape in flat:
        changed = recolor_fill_and_add_border(shape)
        if changed:
            recolor_text_on_yellow_block(shape)
        recolor_line_only(shape)
        recolor_table(shape)
        recolor_chart(shape)
        recolor_blue_text(shape)
    # 노란 fill 셰이프 bbox 확정 후 후처리
    yellow_bboxes = collect_yellow_bboxes(slide)
    blacken_white_text_on_yellow(slide, yellow_bboxes)
    reposition_icons(slide, yellow_bboxes)


def main():
    shutil.copy(SRC, OUT)
    prs = Presentation(OUT)

    # 표지: 사용자가 제공한 간판 클로즈업 사진을 듀오톤으로 가공해 우측 배경
    here = os.path.dirname(os.path.abspath(__file__))
    src_photo = os.path.join(here, "이마트 간판.jpg")
    bg_path = os.path.join(here, "_cover_bg.png")
    sw_in = prs.slide_width / 914400
    sh_in = prs.slide_height / 914400
    make_cover_background(src_photo, bg_path, int(sw_in / 2 * 300), int(sh_in * 300))

    for idx, slide in enumerate(prs.slides, start=1):
        process_slide(slide)
        if idx == 1:
            style_cover_slide(slide, prs, bg_path)
    recolor_all_picture_parts(prs)
    prs.save(OUT)
    print(f"SAVED: {OUT}")
    print(f"SLIDES: {len(prs.slides)}")


if __name__ == "__main__":
    main()
