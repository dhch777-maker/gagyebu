import pytest
from layout import get_layout, BlockSpec


def test_get_layout_returns_dict_with_six_blocks():
    layout = get_layout(1080)
    expected_keys = {"meta_top", "wordmark", "headline", "kor_line1", "kor_line2", "handle"}
    assert set(layout.keys()) == expected_keys


def test_each_block_is_blockspec():
    layout = get_layout(1080)
    for block in layout.values():
        assert isinstance(block, BlockSpec)


def test_meta_top_at_y_60():
    layout = get_layout(1080)
    assert layout["meta_top"].y == 60


def test_wordmark_at_y_180_size_96():
    layout = get_layout(1080)
    assert layout["wordmark"].y == 180
    assert layout["wordmark"].font_size == 96


def test_headline_at_y_720():
    layout = get_layout(1080)
    assert layout["headline"].y == 720


def test_kor_lines_40px_apart():
    layout = get_layout(1080)
    assert layout["kor_line2"].y - layout["kor_line1"].y == 40


def test_handle_at_y_1020():
    layout = get_layout(1080)
    assert layout["handle"].y == 1020


def test_all_blocks_center_aligned():
    layout = get_layout(1080)
    for block in layout.values():
        assert block.align == "center"


def test_unsupported_canvas_raises():
    with pytest.raises(ValueError):
        get_layout(720)
