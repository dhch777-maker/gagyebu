"""배경 이미지에 텍스트를 합성. layout.py의 좌표·폰트 사양을 사용."""

from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

from layout import get_layout, DIVIDER, BlockSpec


_HERE = Path(__file__).parent

# 시스템 폰트 후보 (Windows). Noto Sans KR이 없으면 Malgun Gothic으로 폴백.
_NOTO_BOLD_CANDIDATES = [
    Path("C:/Windows/Fonts/NotoSansKR-Bold.otf"),
    Path("C:/Windows/Fonts/malgunbd.ttf"),
]
_NOTO_REGULAR_CANDIDATES = [
    Path("C:/Windows/Fonts/NotoSansKR-Regular.otf"),
    Path("C:/Windows/Fonts/malgun.ttf"),
]
_MONO_CANDIDATES = [
    Path("C:/Windows/Fonts/consola.ttf"),
    Path("C:/Windows/Fonts/cour.ttf"),
]


def _first_existing(candidates: list[Path]) -> Path:
    for p in candidates:
        if p.exists():
            return p
    raise FileNotFoundError(f"None of these fonts found: {candidates}")


FONT_PATHS = {
    "druk":         _HERE / "fonts" / "Druk-Heavy-Trial.otf",
    "noto_bold":    _first_existing(_NOTO_BOLD_CANDIDATES),
    "noto_regular": _first_existing(_NOTO_REGULAR_CANDIDATES),
    "mono":         _first_existing(_MONO_CANDIDATES),
}


def _hex_to_rgb(hex_str: str) -> tuple[int, int, int]:
    s = hex_str.lstrip("#")
    return tuple(int(s[i:i + 2], 16) for i in (0, 2, 4))  # type: ignore[return-value]


def _load_font(spec: BlockSpec) -> ImageFont.FreeTypeFont:
    return ImageFont.truetype(str(FONT_PATHS[spec.font_family]), size=spec.font_size)


def _draw_centered(draw: ImageDraw.ImageDraw, text: str, spec: BlockSpec, canvas_w: int) -> None:
    font = _load_font(spec)
    # textbbox로 픽셀 정확 가운데 정렬
    bbox = draw.textbbox((0, 0), text, font=font)
    text_w = bbox[2] - bbox[0]
    x = (canvas_w - text_w) // 2 - bbox[0]
    draw.text((x, spec.y), text, font=font, fill=_hex_to_rgb(spec.color))


def compose(
    background: Image.Image,
    texts: dict[str, str],
    draw_divider: bool = True,
) -> Image.Image:
    """배경 위에 텍스트 블록 + (선택) 디바이더 라인을 그려 새 PIL Image 반환.

    texts dict에 없는 layout 블록은 건너뜀.
    draw_divider=False면 디바이더 라인도 생략.
    """
    canvas = background.copy().convert("RGB")
    draw = ImageDraw.Draw(canvas)
    layout = get_layout(canvas.width)

    if draw_divider:
        div_x_start = (canvas.width - DIVIDER["width"]) // 2
        div_x_end = div_x_start + DIVIDER["width"]
        div_color = _hex_to_rgb(DIVIDER["color"])
        draw.rectangle(
            [(div_x_start, DIVIDER["y"]), (div_x_end, DIVIDER["y"] + DIVIDER["thickness"])],
            fill=div_color,
        )

    for key, spec in layout.items():
        if key not in texts:
            continue
        _draw_centered(draw, texts[key], spec, canvas.width)

    return canvas
