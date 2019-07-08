__version__ = '0.0.1'
__VERSION__ = __version__
# print(f'xlmaker version: {__version__}')
from .main import XlWorkbook, get_field, \
    divide as div, multiply as mult
from .templates import wb_factory, Column
