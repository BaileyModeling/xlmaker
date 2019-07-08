# import xlsxwriter
from xlsxwriter.workbook import Workbook
from xlsxwriter.chartsheet import Chartsheet
from xlsxwriter.worksheet import Worksheet, convert_cell_args, \
    convert_range_args, convert_column_args
from xlsxwriter.utility import xl_rowcol_to_cell, \
    xl_cell_to_rowcol, xl_range

"""
Extension of xlsxwriter that allows you to easily
and cleanly augment the formatting of a cell.
Cells can be written and modified in any order.
This is accomplished by storing all cell contents
and styles and then sorting them before writing
them to the workbook.
Rather than creating Format objects, create Style
objects and save them to a StyleSheet. Style objects
are essentially just dictionaries of properties.
You can pass two different StyleSheets to different Worksheets.
So, you can use 'header' style in two different sheets
and get two different results.
"""


def get_field(instance, field):
    field_path = field.split('.')
    attr = instance
    for elem in field_path:
        try:
            attr = getattr(attr, elem)
        except AttributeError:
            return None
    return attr


class Style(object):
    def __init__(self, name, properties):
        self.name = name
        if isinstance(properties, dict):
            self._properties = properties
        else:
            raise TypeError
        self.format = None

    def __str__(self):
        result = self.name + ': \r\n'
        for key, value in self._properties.items():
            result += f'  {key}: {str(value)}; \r\n'
        return result

    def __add__(self, other):
        if self._properties == other._properties:
            return self
        name = self.name + "_" + other.name
        properties = {**self._properties, **other._properties}
        return Style(name, properties)

    def __radd__(self, other):
        if other == 0:
            return self
        else:
            return self.__add__(other)

    def __eq__(self, other):
        return self.__dict__ == other.__dict__

    def includes(self, other):
        '''True if all properties of other included and identical to self.'''
        result = True
        for key, value in other._properties.items():
            if (
                key not in self._properties or
                value != self._properties[key]
            ):
                result = False
                break
        return result

    def add(self, property, value):
        self._properties[property] = value

    def get_properties(self):
        return self._properties

    def build(self, workbook):
        if not self.format:
            self.format = workbook.add_format(self.get_properties())
        # return self.format

    def get_format(self, workbook):
        if not self.format:
            self.format = workbook.add_format(self.get_properties())


class StyleSheet(object):
    def __init__(self, styles=None):
        self._converted = False
        self._styles = {}
        self._default = None
        self._boxes = 0
        if styles is not None:
            for s in styles.items():
                if not isinstance(s, Style):
                    raise TypeError
                self._styles[s.name] = s

    def __getattr__(self, name):
        return self._styles.get(name)

    def get(self, name):
        return self._styles.get(name)

    def add(self, style):
        if not isinstance(style, Style):
            raise TypeError
        self._styles[style.name] = style
        return style

    def create(self, name, properties):
        """ Create will include default properties. """
        if self.get(name):
            raise ValueError(f"Style name '{name}' is already in use.")
        if self._default:
            properties = {**self._default.get_properties(), **properties}
        style = Style(name, properties)
        return self.add(style)

    def default(self, properties):
        style = Style('default', properties)
        d = self.add(style)
        self._default = d
        return d

    def extend(self, base, name, properties):
        if not isinstance(base, Style):
            base = self.get(base)
        props = {**base._properties, **properties}
        style = Style(name, props)
        return self.add(style)

    def load(self, style_dict):
        for name, properties in style_dict.items():
            if name == 'default':
                self.default(properties)
            else:
                self.create(name, properties)
        return self

    def get_styles(self):
        return self._styles.items()

    def build(self, workbook):
        if not self._converted:
            for _, s in self.get_styles():
                s.build(workbook)
        self._converted = True

    def __str__(self):
        result = ''
        for name, style in self.get_styles():
            result += str(style)
            result += '\r\n'
        return result


class Table(object):
    pass


class TColumn(object):
    pass


class Row(object):
    def __init__(self,
        row,
        height,
        style=None,
        options=None):

        self.row = row
        self.height = height
        self.style = style
        self.options = options

    def __eq__(self, other):
        return (self.row == other.row)

    def __ne__(self, other):
        return (self.row != other.row)

    def __lt__(self, other):
        return (self.row < other.row)

    def __le__(self, other):
        return (self.row <= other.row)

    def __gt__(self, other):
        return (self.row > other.row)

    def __ge__(self, other):
        return (self.row >= other.row)

    def get_format(self):
        if self.style:
            return self.style.format
        else:
            return None

    def write(self, sheet):
        fmt = self.get_format()
        sheet.set_row(self.row,
            self.height,
            fmt,
            self.options
            )


class Cell(object):
    def __init__(
        self,
        row,
        col,
        value=None,
        style=None,
        data_type=None,
        kwargs=None
    ):
        #todo: url can have 'string' or 'tip' kwargs

        self.value = value
        self.style = style
        self.data_type = data_type
        self.row = row
        self.col = col
        if kwargs is None:
            self.kwargs = {}
        else:
            self.kwargs = kwargs

    def __str__(self):
        value = str(self.value) if self.value is not None else ''
        style = self.style.name if self.style is not None else ''
        data_type = self.data_type if self.data_type is not None else ''
        result = 'value: ' + value + '\r\n'
        result += 'style: ' + style + '\r\n'
        result += 'data_type: ' + data_type + '\r\n'
        return result

    def __eq__(self, other):
        return ((self.row, self.col) == (other.row, other.col))

    def __ne__(self, other):
        return ((self.row, self.col) != (other.row, other.col))

    def __lt__(self, other):
        return ((self.row, self.col) < (other.row, other.col))

    def __le__(self, other):
        return ((self.row, self.col) <= (other.row, other.col))

    def __gt__(self, other):
        return ((self.row, self.col) > (other.row, other.col))

    def __ge__(self, other):
        return ((self.row, self.col) >= (other.row, other.col))

    def add_style(self, style):
        if self.style is None:
            self.style = style
        elif self.style.name == style.name:
            # check if name is same?
            # print(style.name + ' is already in use.')
            pass
        elif self.style == style:
            # print(style.name + ' is identical to existing style.')
            pass
        elif self.style.includes(style):
            # print(style.name + ' properties already included in ' + self.style.name)
            pass
        elif style.includes(self.style):
            # print(self.style.name + ' properties already included in ' + style.name)
            self.style = style
        else:
            self.style += style
        return self.style

    def set_style(self, style):
        self.style = style
        return self.style

    def set_value(self, value):
        self.value = value

    def set_type(self, data_type):
        self.data_type = data_type

    def get_format(self):
        if self.style:
            return self.style.format
        else:
            return None

    def write(self, sheet):
        fmt = self.get_format()
        if self.value is None or self.value == '':
            return sheet.write_blank(
                self.row, self.col, '', fmt
            )

        if self.data_type == 'number':
            return sheet.write_number(
                self.row,
                self.col,
                self.value,
                fmt
            )
        elif self.data_type == 'str':
            return sheet.write_string(
                self.row,
                self.col,
                self.value,
                fmt
            )
        elif self.data_type == 'datetime':
            return sheet.write_datetime(
                self.row,
                self.col,
                self.value,
                fmt
            )
        elif self.data_type == 'formula':
            return sheet.write_formula(
                self.row,
                self.col,
                self.value,
                fmt,
                **self.kwargs
            )
        elif self.data_type == 'url':
            return sheet.write_url(
                self.row,
                self.col,
                self.value,
                fmt,
                **self.kwargs
            )
        else:
            #todo: call more specific write function
            return sheet.write(
                self.row,
                self.col,
                self.value,
                fmt,
                **self.kwargs
            )


class XlWorksheet(Worksheet):
    def __init__(self, *args, **kwargs):
        super(XlWorksheet, self).__init__(*args, **kwargs)
        self._cells = {}
        self._rows = []
        self._row = 0
        self._col = 0
        self._lastrow = 0
        self._lastcol = 0

    def add_stylesheet(self, stylesheet):
        self._css = stylesheet
        return self._css

    def next_row(self, number=1, reset_column=True):
        self._row += number
        if reset_column:
            self._col = 0
        return self._row

    def next_col(self, number=1):
        self._col += number
        return self._col

    def reset_col(self):
        self._col = 0

    def cell(self, value, style=None, data_type=None, row=None, col=None):
        row = row if row is not None else self._row
        col = col if col is not None else self._col
        self._row = row
        self._col = col + 1
        # if row is None or col is None:
        #     row = self._row
        #     col = self._col
        #     self.next_col()
        location = xl_rowcol_to_cell(row, col)
        if isinstance(style, str):
            style = self._css.get(style)
        if style is None:
            style = self._css._default
        # if the cell already has a style,
        # then we will extend it and the new style
        # will need to be added to the stylesheet.
        if location not in self._cells:
            self._cells[location] = Cell(row, col, value, style, data_type)
        else:
            self._cells[location].set_value(value)
            self._cells[location].add_style(style)
            self._cells[location].set_type(data_type)
        if style is not None:
            self._css.add(self._cells[location].style)

    def xy(self, x_rel=0, y_rel=0, x_abs=False, y_abs=False, abs=False):
        location = xl_rowcol_to_cell(
            self._row + y_rel, self._col + x_rel, y_abs, x_abs
        )
        if abs:
            location = f"'{self.name}'!{location}"
        return location

    def x(self, x_rel=0):
        return self._col + x_rel

    def y(self, y_rel=0):
        return self._row + y_rel

    def loc(self, x=None, y=None, x_abs=False, y_abs=False, abs=False):
        if x is None:
            x = self._col
        if y is None:
            y = self._row
        location = xl_rowcol_to_cell(
            y, x, y_abs, x_abs
        )
        if abs:
            location = f"'{self.name}'!{location}"
        return location

    def set_format(self, row, col, style):
        """Will replace any existing format."""
        if isinstance(style, str):
            style = self._css.get(style)
        #todo: allow A1 notation
        location = xl_rowcol_to_cell(row, col)
        if location in self._cells:
            self._cells[location].set_format(style)
        else:
            self._cells[location] = Cell(row, col, style=style)
        self._css.add(self._cells[location].style)

    def format_cell(self, row, col, style):
        """Will extend any existing format."""
        #todo: allow A1 notation
        style = self.get_style(style)
        location = xl_rowcol_to_cell(row, col)
        if location in self._cells:
            self._cells[location].add_style(style)
        else:
            self._cells[location] = Cell(row, col, style=style)
        self._css.add(self._cells[location].style)

    def format_range(
        self, row1=None, col1=None, row2=None, col2=None, style=None
    ):
        style = self.get_style(style)
        if row1 is None:
            row1 = self._row
        if row2 is None:
            row2 = self._row
        if col1 is None:
            col1 = self._col
        if col2 is None:
            col2 = self._col
        for x in range(row1, row2 + 1):
            for y in range(col1, col2 + 1):
                self.format_cell(x, y, style)

    def get_style(self, style):
        if isinstance(style, str):
            style = self._css.get(style)
        elif isinstance(style, Style):
            # check for existing style
            existing = self._css.get(style.name)
            if existing and existing == style:
                style = existing
            elif existing:
                raise ValueError(
                    f"Different style with same name: {style.name}"
                )
        else:
            raise TypeError
        return style

    def vtotal(
        self, num_rows, style=None, dirn=-1,
        ftype=0, row=None, col=None
    ):
        """dirn is direction up/down """
        if row is None:
            row = self._row
        if col is None:
            col = self._col
        rng = xl_range(row + num_rows * dirn, col, row + 1 * dirn, col)
        formula = self.total_formula(rng, ftype)
        result = self.cell(formula, self.get_style(style),
            'formula', row=row, col=col)
        return result

    def htotal(
        self, num_cols, style=None, dirn=-1,
        ftype=0, row=None, col=None
    ):
        if row is None:
            row = self._row
        if col is None:
            col = self._col
        rng = xl_range(row, col + num_cols * dirn, row, col + 1 * dirn)
        formula = self.total_formula(rng, ftype)
        result = self.cell(
            formula, self.get_style(style),
            'formula', row=row, col=col
        )
        return result

    def total_formula(self, range, ftype=0):
        subtot = ftype==1 or \
            (isinstance(ftype, str) and ftype.lower()=='subtotal')
        if subtot:
            return '=SUBTOTAL(9,' + range + ')'
        else:
            return '=SUM(' + range + ')'

    def div_formula(
        self, r_numer, c_numer, r_denom, c_denom, default=None, neg=False
    ):
        numer = xl_rowcol_to_cell(r_numer, c_numer)
        denom = xl_rowcol_to_cell(r_denom, c_denom)
        sign = ''
        if neg:
            sign = '-'
        if default is None:
            formula = f"={sign}{numer}/{denom}"
        else:
            formula = f"=IFERROR({sign}{numer}/{denom},{str(default)})"
        return formula

    def hdiv_formula(self, numer_x_rel, denom_x_rel, default=None):
        """Divide two cells in the current row given two column positions 
        relative to the current cell.
        """
        return self.div_formula(
            self._row,
            self._col + numer_x_rel,
            self._row,
            self._col + denom_x_rel,
            default
        )

    def vdiv_formula(self, numer_y_rel, denom_y_rel, default=None):
        """Divide two cells in the current column given two row positions 
        relative to the current cell.
        """
        return self.div_formula(
            self._row + numer_y_rel,
            self._col,
            self._row + denom_y_rel,
            self._col,
            default
        )

    def divide(self, numer, denom, default=None, style=None):
        if default is None:
            formula = f"={numer}/{denom}"
        else:
            formula = f"=IFERROR({numer}/{denom},{str(default)})"
        return self.cell(formula, style, 'formula')

    def sumifs(self, sum_range, *criteria, style=None):
        formula = f'=SUMIFS({ sum_range }'
        for c in criteria:
            formula += f', { c[0] }, "{ c[1] }"'
        formula += ')'
        return self.cell(formula, style, 'formula')

    # def get_cell(self, row, col):
    #     location = xl_rowcol_to_cell(row, col)
    #     return self._cells.get(location, None)

    def build(self, workbook):
        self._css.build(workbook)
        for row in sorted(self._rows):
            row.write(self)
        for cell in sorted(self._cells.values()):
            cell.write(self)

    def print_cells(self):
        for loc, cell in self._cells.items():
            print(loc + ': ' + str(cell))

    def row_style(self, height=14.25, style=None, options=None, row=None):
        if row is None:
            row = self._row
            # self.next_row()
        if isinstance(style, str):
            style = self._css.get(style)
        if style is None:
            style = self._css._default
        self._rows.append(Row(row, height, style, options))

    def col_style(self):
        #todo
        pass

    def box(
        self, row_1, col_1, row_2, col_2, border_style=1,
        border_color='black', pattern=0, bg_color=0, fg_color=0
    ):
        """Makes an RxC box. Use integers, not the 'A1' format"""
        rows = row_2 - row_1 + 1
        cols = col_2 - col_1 + 1
        num = self._css._boxes
        self._css._boxes += 1

        for x in range((rows) * (cols)):  # Total cells in the rectangle
            properties = {}   # The format resets each loop
            name = str(num) + '_'
            row = row_1 + (x // cols)
            column = col_1 + (x % cols)

            if x < (cols):                     # If it's on the top row
                properties['top'] = border_style
                name += 't'
            if x >= ((rows * cols) - cols):    # If it's on the bottom row
                properties['bottom'] = border_style
                name += 'b'
            if x % cols == 0:                  # If it's on the left column
                properties['left'] = border_style
                name += 'l'
            if x % cols == (cols - 1):         # If it's on the right column
                properties['right'] = border_style
                name += 'r'

            if pattern:
                properties['pattern'] = pattern
            if bg_color:
                properties['bg_color'] = bg_color
            if fg_color:
                properties['fg_color'] = fg_color
            if properties != {}:
                if border_color:
                    properties['border_color'] = border_color
                self.format_cell(row, column, Style(name, properties))

    def set_col_width(self, width=None, options={}, num_columns=1):
        self.set_column(
            self.current_col,
            self.current_col + num_columns,
            width,
            None,
            options
        )

    def set_col_widths(self, *args):
        for x in range(len(args)):
            self.set_column(x, x, args[x])

    def blank_row(self, height=None, style=None):
        height = height if not None else 14.25
        # self.next_row()
        self.cell(None, style)
        self.row_style(height, style)
        self.next_row()


class XlWorkbook(Workbook):
    chartsheet_class = Chartsheet
    worksheet_class = XlWorksheet

    def __init__(self, *args, **kwargs):
        super(XlWorkbook, self).__init__(*args, **kwargs)
        self._css = StyleSheet()

    def sheet(self, name=None, stylesheet=None):
        """
        Add a new worksheet to the Excel workbook.

        Args:
            name: The worksheet name. Defaults to 'Sheet1', etc.

        Returns:
            Reference to a worksheet object.

        """

        ws = self._add_sheet(name, worksheet_class=self.worksheet_class)
        # set default sheet properties
        ws.set_paper(1)
        ws.center_horizontally()
        ws.hide_gridlines(2)  # hide printed grid lines
        ws.set_margins(left=0.5, right=0.5, top=0.4, bottom=0.6)
        ws.set_zoom(90)
        ws.set_print_scale(85)
        if stylesheet is None:
            ws.add_stylesheet(self._css)
        else:
            ws.add_stylesheet(stylesheet)
        # ws.set_theme(self.theme)
        return ws

    def h_sheet(self, *args, **kwargs):
        ws = self.sheet(*args, **kwargs)
        ws.set_landscape()
        return ws

    def add_stylesheet(self, stylesheet):
        self._css = stylesheet
        return self._css

    def load_styles(self, style_dict):
        return self._css.load(style_dict)

    def build_styles(self):
        if not hasattr(self, '_css'):
            return None
        for name, style in self._css.get_styles():
            self._css.name = self.add_format(style.get_properties())

    def build(self):
        for ws in self.worksheets_objs:
            ws.build(self)
        self.close()

    def get_format(self, name):
        return self._css.get(name).get_format(self)


# ========================================================
# Other Functions
# ========================================================

def divide(numerator, denominator, default=0):
    if numerator is None:
        numerator = 0
    if not denominator:
        result = default
    else:
        result = numerator / denominator
    return result


def multiply(*args):
    result = 1
    for x in args:
        if x:
            result = result * x
        else:
            result = 0
    return result


def main():
    filename = 'example.xlsx'
    options = {
        'constant_memory': True,
        'default_date_format': 'dd/mm/yy',
    }
    wb = XlWorkbook(filename, options)
    wb.formats[0].set_font_size(9)
    wb.formats[0].set_font_name('Century Gothic')
    # wb.add_stylesheet(css)
    ws = wb.sheet('tab1')
    ws.box(1, 1, 5, 5)
    ws.box(2, 3, 6, 6)
    # print(ws._css)
    # ws.print_cells()
    wb.build()


if __name__ == '__main__':
    main()
