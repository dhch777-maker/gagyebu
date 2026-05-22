"""시안별 설정 dict. 신규 시안 추가는 SPECS dict에 entry 한 줄 추가로 끝남."""

SPECS = {
    "03": {
        "slug": "mark-the-start",
        "flux_prompt": (
            "minimalist art poster background, single deep cobalt blue ink "
            "splatter on warm ivory paper, organic dynamic splash shape "
            "positioned in upper-center area, small ink droplets scattered "
            "around the main splash, generous negative space top and bottom, "
            "high contrast, fine paper grain texture, no text no letters, "
            "1:1 square format, editorial poster aesthetic, clean composition"
        ),
        "flux_negative": (
            "text, letters, words, watermark, signature, multiple colors, "
            "busy composition, photograph, realistic person, frame, border"
        ),
        "seed": 20260522,
        "draw_divider": False,
        "texts": {
            "meta_top": "ZERO TO ART  ·  Vol. 03",
            "headline": "MARK THE START.",
            "kor_line1": "미술이 시작되는 순간을",
            "kor_line2": "ZETO ART에서 그려갑니다.",
            "handle": "@ZETO_magok",
        },
    },
}


def get_spec(spec_id: str) -> dict:
    """시안 ID로 설정 dict 조회. 없으면 KeyError."""
    return SPECS[spec_id]
