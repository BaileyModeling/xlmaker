

class Cell(object):
    def __init__(
        self,
        row,
        col,
        value=None,
        style=None,
        data_type=None,
        **kwargs
    ):
        # todo: url can have 'string' or 'tip' kwargs
        self.row = row
        self.col = col
        self.value = value
        self.style = style
        self.data_type = data_type
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

    def print(self):
        print(f"Cell(row={self.row}, col={self.col})")
        print(f"   value: {str(self.value)}")
        print(f"   data_type: {str(self.data_type)}")
        print(f"   style: {self.style.name}")
        # print(f"   style: {str(self.style)}")

    def add_style(self, style):
        # print(f"add_style({self.style.name}, {style.name}):")
        if self.style is None:
            self.style = style
        elif self.style.name == style.name:
            # check if name is same?
            # print(style.name + ' is already in use.')
            pass
        elif self.style == style:
            print(style.name + ' is identical to existing style.')
            pass
        elif self.style.includes(style):
            print(style.name + ' properties already included in ' + self.style.name)
            pass
        elif style.includes(self.style):
            print(self.style.name + ' properties already included in ' + style.name)
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
        if self.style is None:
            style_name = "None"
        else:
            style_name = self.style.name
        # print(f"r{self.row} c{self.col}: {style_name: <20}: {str(fmt)}")
        if self.value is None or self.value == '':
            return sheet.write_blank(self.row, self.col, '', fmt)
        elif self.data_type == 'number':
            return sheet.write_number(self.row, self.col, self.value, fmt)
        elif self.data_type == 'str':
            return sheet.write_string(self.row, self.col, self.value, fmt)
        elif self.data_type == 'datetime':
            return sheet.write_datetime(self.row, self.col, self.value, fmt)
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

