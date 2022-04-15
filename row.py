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
