from xlmaker import StyleSheet


class StyleSheetTemplate(StyleSheet):

    def __init__(
        self,
        default_style=None,
        formats=None,
        font_size=11,
        font_name="Calibri",
    ) -> None:
        super().__init__(default_style, formats, font_size, font_name)
        self.title = self.extend('title', {'font_size': '11'})
        self.title_center = self.extend('title_center', {'align': 'center_across', 'font_size': '11'})
        self.title_date = self.extend('title_date', {'num_format': 'mmmm yyyy;@', 'align': 'left', 'bold': True})
        self.subtitle = self.extend('subtitle', {'font_size': '11'})
        self.footnote = self.extend('footnote', {'font_size': '7'})
        self.tableheader = self.extend('tableheader', {'bg_color': '#366092', 'font_color': 'white', 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
        self.th_no_wrap = self.extend('th_no_wrap', {'text_wrap': False,}, base_style=self.tableheader)
        self.th_left = self.extend('th_left', {'align': 'left',}, base_style=self.tableheader)
        self.th_left_no_wrap = self.extend('th_left_no_wrap', {'text_wrap': False}, base_style=self.th_left)
        self.tableheader_percent = self.extend('tableheader_percent', {'num_format': '0%'}, base_style=self.tableheader)
        self.tablesubheader = self.extend('tablesubheader', {'bg_color': '#D9D9D9', 'align': 'center'})
        self.grey = self.extend('grey', {'bg_color': '#D9D9D9'})
        self.blue = self.extend('blue', {'bg_color': '#DCE6F1'})
        self.green = self.extend('green', {'bg_color': '#ebf1de'})
        self.eq_mult = self.extend('eq_mult', {'num_format': '0.00"x"'})
        self.eqmult = self.extend('eqmult', {}, base_style=self.eq_mult)
        self.acres = self.extend('acres', {'num_format': '#,##0 "Acres"'})
        self.sqft = self.extend('sqft', {'num_format': '#,##0 "SF"'})
        self.lots = self.extend('lots', {'num_format': '#,##0 "Lots"'})
        self.units = self.extend('units', {'num_format': '#,##0 "Units"'})
        self.blue_centered = self.extend('blue_centered', {'bg_color': '#DCE6F1', 'align': 'center_across'})
        self.blue_currency = self.extend('blue_currency', {'bg_color': '#DCE6F1', 'num_format': '$#,##0'})
        self.sf_tableheader = self.extend('sf_tableheader', {'align': 'center', 'valign': 'vcenter', 'text_wrap': True, 'bg_color': '#ebf1de'})
        self.mf_tableheader = self.extend('mf_tableheader', {'align': 'center', 'valign': 'vcenter', 'text_wrap': True, 'bg_color': '#dce6f1'})
        self.com_tableheader = self.extend('com_tableheader', {'align': 'center', 'valign': 'vcenter', 'text_wrap': True, 'bg_color': '#e4dfec'})
        self.mf_center_across = self.extend('mf_center_across', {'align': 'center_across', 'valign': 'vcenter', 'text_wrap': True, 'bg_color': '#dce6f1'})
        self.singlefam = self.extend('singlefam', {'bg_color': '#ebf1de'})
        self.multifam = self.extend('multifam', {'bg_color': '#dce6f1'})
        self.commercial = self.extend('commercial', {'bg_color': '#e4dfec'})
