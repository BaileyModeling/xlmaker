from datetime import date
from xlmaker import XlWorkbook
from xlmaker.examples.example_stylesheet import StyleSheetTemplate


css = StyleSheetTemplate(default_style={'font_size': 9, 'font_name': "Century Gothic"})
wb = XlWorkbook(filename="style_overlap_example2.xlsx", stylesheet=css)
ws = wb.add_worksheet('test1')
ws.set_col_widths(10.9, 10.3)
ws.set_column(11, 11, 10.3)
css.combine_styles(("tableheader", "bold"))
print(css.get("tableheader_bold"))
dt = date(2022, 1, 1)
ws.cell('hello world!', style=["bold", "under"])
ws.cell(dt, row=1, col=1, style='mmmm_yy')

ws.format_range(2, 2, 3, 12, 'tableheader')
cell = ws.get_cell(2, 8)
print(cell.style.name)
ws.format_range(2, 8, 2, 12, css.bold)
cell = ws.get_cell(2, 11)
print(cell.style)
ws.cell("green", row=2, col=8, style=css.green)
cell = ws.string("", row=2, col=9)
ws.cell(dt, row=2, col=10)
ws.cell(dt, row=2, col=11, style='mmmm_yy')
ws.cell("box", row=3, col=9, style=css.box)
ws.cell("grey", row=3, col=10, style='grey')
ws.cell("bold", row=3, col=11, style='bold')
ws.cell(5, row=2, col=2)
style = wb.css.combine_styles(("grey", "eqmult"))
ws.cell(2.2, row=2, col=3, style=style)
print(style)
print(ws.get_cell(3, 8))

wb.build()
