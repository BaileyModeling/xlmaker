from datetime import date
from xlmaker import XlWorkbook
from xlmaker.examples.simple_stylesheet import SimpleStyleSheet


css = SimpleStyleSheet()
print("Beginning StyleSheet:")
css.print()
wb = XlWorkbook(filename="style_overlap_example1.xlsx", stylesheet=css)
ws = wb.add_worksheet('test1')
ws.format_range(1, 0, 5, 5, css.tableheader)
ws.format_range(0, 1, 7, 2, css.grey)
ws.format_range(0, 3, 7, 4, css.date)
cell = ws.cell(999, css.bold, row=1, col=0)
cell = ws.cell(888, row=1, col=1)
cell = ws.cell(777, row=0, col=1)
cell = ws.cell(date(2020,1,1), row=0, col=2)
cell = ws.cell(date(2020,1,1), style=css.mmm_yy, row=1, col=2)
print("===========================")
print("Ending StyleSheet:")
css.print()
print("===========================")
ws.print_cells()
wb.build()
