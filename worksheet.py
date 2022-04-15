from xlsxwriter.worksheet import Worksheet, convert_cell_args, \
    convert_range_args, convert_column_args
from xlsxwriter.utility import xl_rowcol_to_cell, \
    xl_cell_to_rowcol, xl_range
from .cell import Cell
from .style import Style
from .stylesheet import StyleSheet
from .row import Row
from . import errors


class XlWorksheet(Worksheet):
    margins = {'left': 0.5, 'right': 0.5, 'top': 0.4, 'bottom': 0.6}
    zoom = 90
    print_scale = 90
    gridlines = 2  # hide printed grid lines
    centered = True
    paper = 1
    landscape = False
    frozen_rows = 0
    frozen_cols = 0
    col_widths = None
    default_footer = ''

    def __init__(self, stylesheet=None, name=None, workbook=None, footer=None):
        super().__init__()
        self.footer = footer or self.default_footer
        self.css = stylesheet
        self.name = name
        if workbook is not None:
            workbook.add_sheet(self, name)
        self._cells = {}
        self._rows = []
        self._merged = []
        self._row = 0
        self._col = 0
        self._lastrow = 0
        self._lastcol = 0
        self.setup_page_layout()
        self.setup_header()
        self.setup_footer()
        self.setup_body()

    def setup_page_layout(self):
        self.set_paper(self.paper)
        if self.centered:
            self.center_horizontally()
        self.hide_gridlines(self.gridlines)
        self.set_margins(**self.margins)
        self.set_print_scale(self.print_scale)
        self.set_zoom(self.zoom)
        if self.landscape:
            self.set_landscape()
        if self.frozen_rows or self.frozen_cols:
            self.freeze_panes(row=self.frozen_rows, col=self.frozen_cols)
        if self.col_widths:
            self.set_col_widths(*self.col_widths)
            self.num_cols = len(self.col_widths)

    def setup_footer(self):
        if self.footer:
            self.set_footer(self.footer)

    def setup_header(self):
        pass

    def setup_body(self):
        pass

    def set_stylesheet(self, stylesheet: StyleSheet):
        self.css = stylesheet
        return self.css

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
    
    def get_cell(self, row:int, col:int) -> Cell:
        location = xl_rowcol_to_cell(row, col)
        return self._cells.get(location)

    def string(self, value, style=None, row=None, col=None):
        return self.cell(value, style=style, data_type="str", row=row, col=col)

    def number(self, value, style=None, row=None, col=None):
        return self.cell(value, style=style, data_type="number", row=row, col=col)

    def cell(self, value, style=None, data_type=None, row=None, col=None):
        row = row if row is not None else self._row
        col = col if col is not None else self._col
        self._row = row
        self._col = col + 1
        location = xl_rowcol_to_cell(row, col)

        style = self.get_style(style)
        if hasattr(value, "value"):
            value = value.value

        cell = self._cells.get(location)
        if cell is None:
            cell = Cell(row, col, value, style, data_type)
            self._cells[location] = cell
        else:
            cell.set_value(value)
            cell.set_type(data_type)
            existing_style = cell.style
            if existing_style is None:
                cell.style = style
            else:
                combined_style = existing_style + style
                existing_combined_style = self.css.get(combined_style.name)
                if existing_combined_style is None:
                    cell.add_style(style)
                else:
                    cell.style = existing_combined_style
        self.css.add(cell.style, exists_ok=True)
        return cell
        # if location not in self._cells:
        #     self._cells[location] = Cell(row, col, value, style, data_type)
        # else:
        #     self._cells[location].set_value(value)
        #     self._cells[location].set_type(data_type)
        #     if style is not None:
        #         self._cells[location].add_style(style)
        # if style is not None:
        #     self.css.add(self._cells[location].style, exists_ok=True)
        # return self._cells[location]

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

    def xlrange(self, x1=None, y1=None, x2=None, y2=None, dx=0, dy=0):
        if x1 is None:
            x1 = self._col
        if y1 is None:
            y1 = self._row

        if x2 is None:
            x2 = x1 + dx
        if y2 is None:
            y2 = y1 + dy

        return xl_range(y1, x1, y2, x2)

    def set_format(self, row, col, style):
        """Will replace any existing format."""
        if isinstance(style, str):
            style = self.css.get(style)
        location = xl_rowcol_to_cell(row, col)
        if location in self._cells:
            self._cells[location].set_format(style)
        else:
            self._cells[location] = Cell(row, col, style=style)
        self.css.add(self._cells[location].style)

    def format_cell(self, row, col, style):
        """Will extend any existing format."""
        #todo: allow A1 notation
        style = self.get_style(style)
        location = xl_rowcol_to_cell(row, col)
        if location in self._cells:
            cell = self._cells[location]
            existing_style = cell.style
            if existing_style is None:
                cell.style = style
            else:
                combined_style = existing_style + style
                existing_combined_style = self.css.get(combined_style.name)
                if existing_combined_style is None:
                    cell.add_style(style)
                else:
                    cell.style = existing_combined_style
        else:
            cell = Cell(row, col, style=style)
            self._cells[location] = cell
        self.css.add(cell.style, exists_ok=True)

    def format_range(
        self, row1=None, col1=None, row2=None, col2=None, style=None
    ):
        """row and col counts from 0. eg row1=2 is 3rd row of excel sheet."""
        # style = self.get_style(style)
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

    def get_style(self, style) -> Style:
        if isinstance(style, str):
            style = self.css.get(style)
        elif isinstance(style, (list, tuple)):
            style = self.css.combine_styles(style)
        elif isinstance(style, Style):
            style = self.css.add(style, exists_ok=True)
        else:
            style = self.css.default
        # if isinstance(style, str):
        #     style = self.css.get(style)
        # elif isinstance(style, Style):
        #     # check for existing style
        #     existing = self.css.get(style.name)
        #     if existing and existing == style:
        #         style = existing
        #     elif existing:
        #         raise ValueError(
        #             f"Different style with same name: {style.name}"
        #         )
        # else:
        #     raise TypeError
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
        formula = ''
        if num_cols > 0:
            rng = xl_range(row, col + num_cols * dirn, row, col + 1 * dirn)
            formula = self.total_formula(rng, ftype)
        result = self.cell(
            formula, self.get_style(style),
            'formula', row=row, col=col
        )
        return result

    def hvariance(self, c1_rel, c2_rel, style=None):
        row = self.y()
        c1 = self.x() + c1_rel
        c2 = self.x() + c2_rel
        a = xl_rowcol_to_cell(row, c1)
        b = xl_rowcol_to_cell(row, c2)
        return self.cell(f"={a}-{b}", self.get_style(style))

    def total_formula(self, range, ftype=0):
        subtot = ftype==1 or \
            (isinstance(ftype, str) and ftype.lower()=='subtotal')
        if subtot:
            return '=SUBTOTAL(9,' + range + ')'
        else:
            return '=SUM(' + range + ')'

    def mult_formula(self, r1, c1, r2, c2, neg=False):
        a = xl_rowcol_to_cell(r1, c1)
        b = xl_rowcol_to_cell(r2, c2)
        sign = ''
        if neg:
            sign = '-'
        return f"={sign}{a}*{b}"

    def vmult_formula(self, r1_rel, r2_rel, neg=False):
        return self.mult_formula(
            self._row + r1_rel,
            self._col,
            self._row + r2_rel,
            self._col,
            neg
        )

    def hmult_formula(self, c1_rel, c2_rel, neg=False):
        return self.mult_formula(
            self._row,
            self._col + c1_rel,
            self._row,
            self._col + c2_rel,
            neg
        )

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
        self.css.build(workbook)
        for row in sorted(self._rows):
            row.write(self)
        for cell in sorted(self._cells.values()):
            cell.write(self)
        for rng in self._merged:
            pass

    def print_cells(self):
        print(f"Cells for Worksheet {self.name}:")
        for loc in sorted(self._cells):
            print(loc + ': ' + str(self._cells[loc]))

    def row_style(self, height=14.25, style=None, options=None, row=None):
        if row is None:
            row = self._row
            # self.next_row()
        if isinstance(style, str):
            style = self.css.get(style)
        if style is None:
            style = self.css._default
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
        num = self.css._boxes
        self.css._boxes += 1

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
