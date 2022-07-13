from pathlib import Path
from xlsxwriter.workbook import Workbook
from xlsxwriter.chartsheet import Chartsheet
from .worksheet import XlWorksheet
from .stylesheet import StyleSheet
from . import errors


class XlWorkbook(Workbook):
    chartsheet_class = Chartsheet
    worksheet_class = XlWorksheet
    stylesheet_class = StyleSheet
    default_opt = {
        'constant_memory': True,
        'default_date_format': 'dd/mm/yy',
    }
    default_props = {
        'author': 'xlmaker',
        'comments': 'Created with xlmaker'
    }

    def __init__(
        self,
        filename="report.xlsx",
        options=None,
        properties=None,
        stylesheet=None,
    ):
        if filename[-5:] != '.xlsx':
            filename = filename + '.xlsx'
        options = options or {}
        options = {**self.default_opt, **options}
        super().__init__(filename=filename, options=options)
        properties = properties or {}
        properties = {**self.default_props, **properties}
        self.set_properties(properties)
        if isinstance(stylesheet, StyleSheet):
            self.css = stylesheet
        else:
            self.css = self.stylesheet_class()
        self.formats[0].set_font_size(self.css.default.font_size)
        self.formats[0].set_font_name(self.css.default.font_name)

    def add_worksheet(self, name=None, worksheet_class=None):
        worksheet_class = worksheet_class or self.worksheet_class
        ws = self._add_sheet(name, worksheet_class=worksheet_class)
        # set default sheet properties
        ws.set_paper(1)
        ws.center_horizontally()
        ws.hide_gridlines(2)  # hide printed grid lines
        ws.set_margins(left=0.5, right=0.5, top=0.4, bottom=0.6)
        ws.set_zoom(90)
        ws.set_print_scale(85)
        ws.set_stylesheet(self.css)
        return ws

    def add_sheet(self, worksheet:XlWorksheet, name=None):
        name = name or worksheet.name
        sheet_index = len(self.worksheets_objs)
        name = self._check_sheetname(name, isinstance(worksheet, Chartsheet))
        if not hasattr(self, 'max_url_length'):
            self.max_url_length = 255

        # Initialization data to pass to the worksheet.
        init_data = {
            'name': name,
            'index': sheet_index,
            'str_table': self.str_table,
            'worksheet_meta': self.worksheet_meta,
            'constant_memory': self.constant_memory,
            'tmpdir': self.tmpdir,
            'date_1904': self.date_1904,
            'strings_to_numbers': self.strings_to_numbers,
            'strings_to_formulas': self.strings_to_formulas,
            'strings_to_urls': self.strings_to_urls,
            'nan_inf_to_errors': self.nan_inf_to_errors,
            'default_date_format': self.default_date_format,
            'default_url_format': self.default_url_format,
            'excel2003_style': self.excel2003_style,
            'remove_timezone': self.remove_timezone,
            'max_url_length': self.max_url_length,
        }

        worksheet._initialize(init_data)

        self.worksheets_objs.append(worksheet)
        self.sheetnames[name] = worksheet
        worksheet.set_stylesheet(self.css)
        return worksheet

    def h_sheet(self, *args, **kwargs):
        ws = self.sheet(*args, **kwargs)
        ws.set_landscape()
        return ws

    def add_stylesheet(self, stylesheet):
        self.css = stylesheet
        return self.css

    def load_styles(self, style_dict):
        return self.css.load_styles(style_dict)

    def add_style(self, style):
        return self.css.add(style)

    def build_styles(self):
        if not hasattr(self, 'css'):
            return None
        for name, style in self.css.get_styles():
            self.css.name = self.add_format(style.get_properties())

    def build(self):
        filepath = Path(self.filename).resolve().parent
        filepath.mkdir(parents=True, exist_ok=True)
        for ws in self.worksheets_objs:
            ws.build(self)
        self.close()
        print("File created: ", Path(self.filename).resolve())

    def get_format(self, name):
        return self.css.get(name).get_format(self)
