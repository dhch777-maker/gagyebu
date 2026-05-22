import pytest
from prompts import get_spec, SPECS


def test_spec_03_exists():
    assert "03" in SPECS


def test_get_spec_returns_dict_for_known_id():
    spec = get_spec("03")
    assert isinstance(spec, dict)


def test_get_spec_has_required_keys():
    spec = get_spec("03")
    required = {"slug", "flux_prompt", "flux_negative", "seed", "texts"}
    assert required.issubset(spec.keys())


def test_get_spec_texts_has_all_blocks():
    spec = get_spec("03")
    text_keys = {"meta_top", "wordmark", "headline", "kor_line1", "kor_line2", "handle"}
    assert text_keys.issubset(spec["texts"].keys())


def test_get_spec_03_slug_is_mark_the_start():
    assert get_spec("03")["slug"] == "mark-the-start"


def test_get_spec_unknown_id_raises():
    with pytest.raises(KeyError):
        get_spec("99")
