import pytest
from PIL import Image
from compositor import compose, FONT_PATHS


@pytest.fixture
def synthetic_bg():
    """1080² 아이보리 단색 배경 (FLUX 호출 회피)."""
    return Image.new("RGB", (1080, 1080), (244, 241, 234))  # #f4f1ea


@pytest.fixture
def sample_texts():
    return {
        "meta_top": "ZERO TO ART  ·  Vol. 03",
        "wordmark": "ZETO ART",
        "headline": "MARK THE START.",
        "kor_line1": "미술이 시작되는 순간을",
        "kor_line2": "ZETO ART에서 그려갑니다.",
        "handle": "@ZETO_magok",
    }


def test_font_paths_resolved():
    """폰트 파일이 실제로 존재하는지 검증."""
    for key, path in FONT_PATHS.items():
        assert path.exists(), f"Missing font: {key} @ {path}"


def test_compose_returns_pil_image(synthetic_bg, sample_texts):
    result = compose(synthetic_bg, sample_texts)
    assert isinstance(result, Image.Image)


def test_compose_preserves_canvas_size(synthetic_bg, sample_texts):
    result = compose(synthetic_bg, sample_texts)
    assert result.size == (1080, 1080)


def test_compose_actually_draws_text(synthetic_bg, sample_texts):
    """텍스트 그려진 영역과 빈 배경의 픽셀 차이로 검증."""
    result = compose(synthetic_bg, sample_texts)
    # 워드마크는 y=180~276, 가운데 정렬이므로 캔버스 중앙(540, 220) 부근 픽셀이 변해야 함
    center_pixel = result.getpixel((540, 220))
    bg_pixel = synthetic_bg.getpixel((540, 220))
    assert center_pixel != bg_pixel, "워드마크 영역 픽셀이 그대로면 텍스트가 안 그려진 것"


def test_compose_draws_divider_line(synthetic_bg, sample_texts):
    """y=300 위치에 가로 라인(블루)이 그려졌는지 픽셀로 검증."""
    result = compose(synthetic_bg, sample_texts)
    # 디바이더는 가로 480px, 가운데 정렬, y=300, 두께 2 → (540, 300) 부근이 블루여야 함
    px = result.getpixel((540, 300))
    # #266a9c = (38, 106, 156). 정확 비교는 PIL 렌더링 anti-alias 때문에 어려우니 R<G<B 패턴만 확인
    r, g, b = px[:3]
    assert b > r and b > g, f"디바이더 픽셀이 블루가 아님: {px}"
