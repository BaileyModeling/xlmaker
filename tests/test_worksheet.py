import pytest
from datetime import date
from xlmaker.workbook import XlWorkbook
from xlmaker.examples.simple_stylesheet import SimpleStyleSheet


@pytest.fixture
def wb_ws_css():
    css = SimpleStyleSheet()
    wb = XlWorkbook(filename="tests/report.xlsx", stylesheet=css)
    ws = wb.add_worksheet('test1')
    return wb, ws, css


def test_overlap_format_range_has_same_format_object(wb_ws_css):
    wb, ws, css = wb_ws_css
    ws.format_range(1, 0, 5, 5, css.tableheader)
    ws.format_range(0, 1, 7, 2, css.grey)
    wb.build()
    assert ws.get_cell(1, 1).style.format is ws.get_cell(5, 2).style.format
