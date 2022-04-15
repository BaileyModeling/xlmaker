import pytest
from xlmaker import StyleSheet


@pytest.fixture
def css():
    stylesheet = StyleSheet({"font_size": "9", "font_name": "Century Gothic"})
    return stylesheet


def test_stylesheet_init(css):
    assert css.default.get_property("font_size") == "9"
    assert css.default.get_property("font_name") == "Century Gothic"


def test_stylesheet_create_style(css):
    css.create('test_centered', {'align': 'center'})
    assert css.test_centered.get_property("align") == "center"
    assert css.test_centered.get_property("font_size", None) is None


def test_stylesheet_extend_style(css):
    css.extend('test_centered', {'align': 'center'})
    assert css.test_centered.get_property("align") == "center"
    assert css.test_centered.get_property("font_size", None) == "9"


def test_stylesheet_load_styles(css):
    css.load_styles({
        'test_centered': {'align': 'center'},
        'test_right': {'align': 'right'},        
    })
    assert css.test_centered.get_property("align") == "center"
    assert css.test_centered.get_property("font_size", None) == "9"
    assert css.test_right.get_property("align") == "right"
    assert css.test_right.get_property("font_name") == "Century Gothic"


def test_cannot_add_duplicate_name(css):
    with pytest.raises(Exception) as e_info:
        css.create('default', {'align': 'center'})
