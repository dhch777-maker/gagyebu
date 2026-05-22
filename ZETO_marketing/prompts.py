"""시안별 설정 dict. 신규 시안 추가는 SPECS dict에 entry 한 줄 추가로 끝남.

각 entry의 필수 키:
  - slug:          파일명에 들어갈 슬러그
  - flux_prompt:   FLUX 생성 프롬프트
  - flux_negative: negative prompt
  - seed:          재현성용 시드
  - texts:         {meta_top, headline, kor_line1, kor_line2, handle} (5개 필수)

선택 키 (생략 가능):
  - texts["wordmark"]:  큰 ZETO ART 워드마크. FLUX 배경이 y=180~276 영역을 침범하면
                        omit해서 텍스트-배경 충돌을 피한다.
  - draw_divider:       y=300 가로 라인 표시 여부. 기본 True. 워드마크와 같이 끄는 게
                        일반적(워드마크 아래 장식이므로).

신규 슬롯 추가 절차:
  1) FLUX 프롬프트 + seed 결정 → SPECS dict에 entry 추가
  2) `python generate.py NN` 실행 → _bg/NN-bg.png 생성됨
  3) 결과 PNG 확인 → 텍스트 영역과 배경이 충돌하면 wordmark omit 또는
     draw_divider=False로 조정 후 generate.py 재실행 (FLUX 재호출 없음, 캐시 사용)
"""

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
