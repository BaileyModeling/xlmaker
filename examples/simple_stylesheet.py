from xlmaker import StyleSheet


class SimpleStyleSheet(StyleSheet):

    def __init__(
        self,
        font_size=9,
        font_name="Century Gothic",
    ) -> None:
        self._styles = {}
        self._converted = False
        self.default = self.create('default', {'font_size': font_size, 'font_name': font_name})
        self.date = self.extend('date', {'num_format': 'm/d/yy'})
        self.bold = self.extend('bold', {'bold': True})
        self.tableheader = self.extend('tableheader', {'bg_color': '#366092', 'font_color': 'white', 'align': 'center'})
        self.grey = self.extend('grey', {'bg_color': '#D9D9D9'})
        self.mmm_yy = self.extend('mmm_yy', {'num_format': 'mmm yy;@'})
