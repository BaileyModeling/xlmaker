import pytest
from xlmaker import Style


@pytest.fixture
def original():
    style = Style("original", {"font_size": "9", "font_name": "Century Gothic"})
    return style


def test_style_copy(original):
    # original = Style("Original", {'font_size': '9', 'font_name': 'Century Gothic'})
    cp = original.copy()
    cp.set_property("font_size", "11")
    assert cp.get_property("font_size") == "11"
    assert original.get_property("font_size") == "9"


def test_style_append(original):
    original.append(align="center")
    assert original.get_property("font_size") == "9"
    assert original.get_property("align") == "center"


def test_style_append_overwrites_existing(original):
    original.append(align="center", font_size="11")
    assert original.get_property("font_size") == "11"
    assert original.get_property("align") == "center"


def test_style_includes(original):
    cp = original.copy()
    cp.append(align="center")
    assert cp.includes(original)


def test_style_does_not_include(original):
    cp = original.copy()
    original.append(align="center")
    assert not cp.includes(original)


def test_add_identical_styles_returns_self(original):
    cp = original.copy()
    cp.name = "copy"
    result = original + cp
    assert result is original


def test_add_styles_no_intersection(original):
    other = Style("other", {'align': 'center', 'valign': 'vcenter'})
    result = original + other
    assert result.name == "original_other"
    assert result.get_property("align") == "center"
    assert result.get_property("valign") == "vcenter"
    assert result.get_property("font_size") == "9"
    assert result.get_property("font_name") == "Century Gothic"


def test_add_styles_with_intersection(original):
    other = Style("other", {'font_size': '11', 'valign': 'vcenter'})
    result = original + other
    assert result.get_property("font_size") == "11"


def test_add_styles_with_intersection_reverse_order(original):
    other = Style("other", {'font_size': '11', 'valign': 'vcenter'})
    result = other + original
    assert result.get_property("font_size") == "9"
