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
import sys
import shutil
from io import BytesIO
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Pt
from PIL import Image

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


def main():
    shutil.copy(SRC, OUT)
    prs = Presentation(OUT)
    for idx, slide in enumerate(prs.slides, start=1):
        process_slide(slide)
    recolor_all_picture_parts(prs)
    prs.save(OUT)
    print(f"SAVED: {OUT}")
    print(f"SLIDES: {len(prs.slides)}")


if __name__ == "__main__":
    main()
