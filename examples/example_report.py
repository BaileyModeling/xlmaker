from datetime import date
from xlmaker.workbook import XlWorkbook
from xlmaker.worksheet import XlWorksheet
from xlmaker.examples.example_stylesheet import StyleSheetTemplate


"""
The problem with this method is that ExampleSheet doesn't have access to css
when building the class. So there are no hints from the editor. 
"""

class ExampleSheet(XlWorksheet):
    frozen_rows = 0
    frozen_cols = 0
    col_widths = None
    default_footer = f'&R&7{__file__}'

    def __init__(self, wb, name=None):
        super().__init__(stylesheet=wb.css, name=name)
        wb.add_sheet(self, name)

    def setup_header(self):
        self.cell('Page Title', style=["bold", "under"])
        self.next_row(2)

    def setup_body(self):
        self.cell("Page Text")


if __name__=="__main__":
    css = StyleSheetTemplate()
    wb = XlWorkbook(filename="example_report.xlsx", stylesheet=css)
    ws = ExampleSheet(wb, "report1")
    wb.build()
