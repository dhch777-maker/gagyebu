"""텍스트 블록 좌표/폰트 사양. 스펙 섹션 3 레이아웃을 코드로 표현."""

from dataclasses import dataclass


@dataclass(frozen=True)
class BlockSpec:
    y: int
    font_size: int
    font_family: str  # "druk" | "noto_bold" | "noto_regular" | "mono"
    color: str        # hex like "#266a9c"
    align: str        # "center"


# 색상 토큰 (스펙 섹션 3)
COLOR_BLUE = "#266a9c"
COLOR_DARK = "#1a1a1a"
COLOR_MUTE = "#9a978f"


_LAYOUT_1080 = {
    "meta_top": BlockSpec(y=60,   font_size=18, font_family="mono",          color=COLOR_MUTE, align="center"),
    "wordmark": BlockSpec(y=180,  font_size=96, font_family="druk",          color=COLOR_BLUE, align="center"),
    "headline": BlockSpec(y=720,  font_size=72, font_family="druk",          color=COLOR_DARK, align="center"),
    "kor_line1": BlockSpec(y=850, font_size=26, font_family="noto_bold",     color=COLOR_DARK, align="center"),
    "kor_line2": BlockSpec(y=890, font_size=26, font_family="noto_bold",     color=COLOR_DARK, align="center"),
    "handle":   BlockSpec(y=1020, font_size=20, font_family="noto_regular",  color=COLOR_MUTE, align="center"),
}


# 가로 라인 (워드마크 아래)
DIVIDER = {
    "y": 300,
    "width": 480,
    "thickness": 2,
    "color": COLOR_BLUE,
}


def get_layout(canvas_size: int) -> dict[str, BlockSpec]:
    """캔버스 크기에 맞는 레이아웃 반환. 현재는 1080만 지원."""
    if canvas_size != 1080:
        raise ValueError(f"Unsupported canvas size: {canvas_size}. Only 1080 supported.")
    return _LAYOUT_1080
