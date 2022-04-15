from . import errors
from .style import Style


class StyleSheet:

    def __init__(
        self,
        default_style=None,
        formats=None,
        font_size=11,
        font_name="Calibri",
    ) -> None:
        self._styles = {}
        formats = formats or {}
        if default_style is None:
            if "default" in formats:
                default_style = formats.pop("default")
            else:
                default_style = {
                    'font_size': font_size, 'font_name': font_name
                }
        self.default = self.create('default', default_style)
        self.load_styles(formats)
        self._converted = False
        self.setup_standard_styles()

    def setup_standard_styles(self):
        self.spacer = self.extend('spacer', {'font_size': '3'})
        self.centered = self.extend('centered', {'align': 'center'})
        self.left = self.extend('left', {'align': 'justify'})
        self.right = self.extend('right', {'align': 'right'})
        self.vcenter = self.extend('vcenter', {'align': 'vcenter'})
        self.middle = self.extend('middle', {'align': 'center', 'valign': 'vcenter'})
        self.top = self.extend('top', {'align': 'top'})
        self.longtext = self.extend('longtext', {'align': 'center_across', 'text_wrap': True})
        self.longlist = self.extend('longlist', {'align': 'left', 'text_wrap': False, 'align': 'top'})
        self.wrap = self.extend('wrap', {'text_wrap': True})
        self.indent = self.extend('indent', {'indent': 1})
        self.indent1 = self.extend('indent1', {'indent': 1})
        self.indent2 = self.extend('indent2', {'indent': 2})
        self.indent3 = self.extend('indent3', {'indent': 3})
        self.indent4 = self.extend('indent4', {'indent': 4})
        self.indent5 = self.extend('indent5', {'indent': 5})
        self.indent6 = self.extend('indent6', {'indent': 6})
        self.currency = self.extend('currency', {'num_format': '$#,##0'})
        self.currency2 = self.extend('currency2', {'num_format': '$#,##0.00'})
        self.accounting = self.extend('accounting', {'num_format': '_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'})
        self.accounting2 = self.extend('accounting2', {'num_format': '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'})
        self.number = self.extend('number', {'num_format': '#,##0'})
        self.number2 = self.extend('number2', {'num_format': '#,##0.00'})
        self.number3 = self.extend('number3', {'num_format': '#,##0.000'})
        self.percent = self.extend('percent', {'num_format': '0%'})
        self.percent1 = self.extend('percent1', {'num_format': '0.0%'})
        self.percent2 = self.extend('percent2', {'num_format': '0.00%'})
        self.date = self.extend('date', {'num_format': 'm/d/yy'})
        self.m_d_yyyy = self.extend('m_d_yyyy', {'num_format': 'm/d/yyyy;@'})
        self.mm_dd_yyyy = self.extend('mm_dd_yyyy', {'num_format': 'mm/dd/yyyy;@'})
        self.mmm_yyyy = self.extend('mmm_yyyy', {'num_format': 'mmm yyyy;@'})
        self.mmm_yy = self.extend('mmm_yy', {'num_format': 'mmm yy;@'})
        self.mmmm_yyyy = self.extend('mmmm_yyyy', {'num_format': 'mmmm yyyy;@'})
        self.mmmm_yy = self.extend('mmmm_yy', {'num_format': 'mmmm-yy;@'})
        self.bold = self.extend('bold', {'bold': True})
        self.under = self.extend('under', {'underline': True})
        self.number_list = self.extend('number_list', {'num_format': '0.', 'align': 'top'})
        self.center_across = self.extend('center_across', {'align': 'center_across'})
        self.subtotal = self.extend('subtotal', {'top': 1})
        self.total = self.extend('total', {'top': 1})
        self.grandtotal = self.extend('grandtotal', {'top': 1, 'bottom': 6})
        self.months = self.extend('months', {'num_format': '#,##0 "Months"'})
        self.years = self.extend('years', {'num_format': '#,##0 "Years"'})
        self.yesno = self.extend('yesno', {'num_format': '"Yes";-;"No"'})
        self.thousands = self.extend('thousands', {'num_format': '_(* #,##0,_);_(* (#,##0,);_(* "-"??_);_(@_)'})
        self.ol_center_r = self.extend('ol_center_r', {'align': 'center_across', 'right': 1, 'top': 1, 'bottom': 1})
        self.ol_center_l = self.extend('ol_center_l', {'align': 'center_across', 'left': 1, 'top': 1, 'bottom': 1})
        self.ol_center_mid = self.extend('ol_center_mid', {'align': 'center_across', 'top': 1, 'bottom': 1})
        self.box = self.extend('box', {'top': 1, 'bottom': 1, 'right': 1, 'left': 1, 'align': 'center'})
        self.acco_thousands = self.extend('acco_thousands', {'num_format': '_($* #,##0,_);_($* (#,##0,);_($* "-"??_);_(@_)'})

    def __str__(self):
        result = ''
        for name, style in self._styles.items():
            result += str(style)
            result += '\r\n'
        return result

    def __getattr__(self, name):
        return self._styles.get(name)

    def get(self, name):
        return self._styles.get(name)

    def get_styles(self):
        return self._styles.items()

    def add(self, style, exists_ok=False):
        if not isinstance(style, Style):
            raise TypeError
        if not exists_ok and style.name in self._styles:
            raise errors.DuplicateKeyError(f"{style.name} is already in use.")
        self._styles[style.name] = style
        return style

    def create(self, name, properties):
        if self.get(name):
            raise ValueError(f"Style name '{name}' is already in use.")
        style = Style(name, properties)
        return self.add(style)

    def extend(self, name, properties, base_style='default'):
        if not isinstance(base_style, Style):
            base_style = self.get(base_style)
        style = base_style.extend(name, properties)
        return self.add(style)

    def load_styles(self, style_dict):
        for name, properties in style_dict.items():
            self.extend(name, properties)
        return self

    def combine_styles(self, styles) -> Style:
        """styles: list of string names of styles"""
        objects = []
        # for style in sorted(styles): sorting impacts cascade
        for style in styles:
            obj = self.get(style)
            if obj is None:
                raise ValueError(f"Style name does not exist: {style} ")
            objects.append(obj)
        combined = sum(objects)
        self.add(combined, exists_ok=True)
        return combined

    def name_in_use(self, name):
        return name in self._styles

    def build(self, workbook):
        if not self._converted:
            for s in self._styles.values():
                s.build(workbook)
            self._converted = True

    def print(self):
        for name, style in self._styles.items():
            print(style)

