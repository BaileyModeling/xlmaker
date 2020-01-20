from . import XlWorkbook, get_field, div, mult
from collections import namedtuple


Column = namedtuple('Column', 'name attr width fmt')
Column.__new__.__defaults__ = (8.43, 'default')


formats = {
    'default': {'font_size': '9', 'font_name': 'Century Gothic'},
    'spacer': {'font_size': '3'},
    'centered': {'align': 'center'},
    'left': {'align': 'justify'},
    'right': {'align': 'right'},
    'longtext': {'align': 'center_across', 'text_wrap': True},
    'indent': {'indent': 1},
    'indent2': {'indent': 2},
    'indent3': {'indent': 3},
    'currency': {'num_format': '$#,##0'},
    'currency2': {'num_format': '$#,##0.00'},
    'accounting': {'num_format': '_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'},
    'accounting2': {'num_format': '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)'},
    'number': {'num_format': '#,##0'},
    'number3': {'num_format': '#,##0.000'},
    'percent': {'num_format': '0%'},
    'percent1': {'num_format': '0.0%'},
    'percent2': {'num_format': '0.00%'},
    'date': {'num_format': 'm/d/yy'},
    'blue': {'bg_color': '#DCE6F1'},
    'grey': {'bg_color': '#D9D9D9'},
    'bold': {'bold': True},
    'eq_mult': {'num_format': '0.00"x"'},
    'eqmult': {'num_format': '0.00"x"'},
    'center_across': {'align': 'center_across'},
    'title': {'font_size': '11'},
    'subtitle': {'font_size': '11'},
    'footnote': {'font_size': '7'},
    'tableheader': {'bg_color': '#366092', 'font_color': 'white', 'align': 'center'},  # vcenter
    'tablesubheader': {'bg_color': '#D9D9D9', 'align': 'center'},  # vcenter
    'subtotal': {'top': 1},
    'total': {'top': 1},
    'grandtotal': {'top': 1, 'bottom': 6},  # double bottom
}


def wb_factory(filename, options=None, properties=None):
    default_opt = {
        'constant_memory': True,
        'default_date_format': 'dd/mm/yy',
        # 'default_format_properties': default.get_properties()
    }
    if options is None:
        options = default_opt
    elif isinstance(options, dict):
        options = {**default_opt, **options}
    else:
        raise TypeError

    wb = XlWorkbook(filename, options)

    default_props = {
        'title': 'Stratus',
        'subject': 'Stratus',
        'author': 'Bailey Edwards',
        'company': 'Stratus Properties',
        'keywords': 'Stratus',
        'comments': 'Created with Python and XlsxWriter'}
    if properties is None:
        properties = default_props
    elif isinstance(properties, dict):
        properties = {**default_props, **properties}
    else:
        raise TypeError
    wb.set_properties(properties)

    wb.formats[0].set_font_size(9)
    wb.formats[0].set_font_name('Century Gothic')
    css = wb.load_styles(formats)
    css.add(css.total + css.currency)
    css.add(css.total + css.currency2)
    css.add(css.total + css.accounting)
    css.add(css.total + css.accounting2)
    css.add(css.subtotal + css.accounting2)
    css.add(css.total + css.number)
    css.add(css.total + css.number3)
    css.add(css.total + css.percent)
    css.add(css.total + css.percent1)
    css.add(css.grandtotal + css.currency2)
    css.add(css.grandtotal + css.accounting2)
    css.add(css.grey + css.bold)
    css.add(css.grey_bold + css.right)
    css.add(css.grey + css.currency)
    css.add(css.grey + css.currency2)
    css.add(css.grey + css.number)
    css.add(css.grey + css.number3)
    css.add(css.grey + css.percent1)
    css.add(css.grey + css.date)
    css.add(css.grey_number + css.total)
    css.add(css.grey_number3 + css.total)
    css.add(css.grey_percent1 + css.total)
    css.add(css.tableheader + css.date)

    return wb
