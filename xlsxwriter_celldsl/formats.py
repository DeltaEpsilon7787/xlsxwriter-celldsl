from collections import defaultdict
from typing import Any, Dict

from attr import Factory, attrs
from xlsxwriter import Workbook as XlsxWriterWorkbook
from xlsxwriter.format import Format


class FormatDict(Dict[str, Any]):
    """A special variant of vanilla dictionary that implement __or__ and __hash__. Used to create and merge formats.

    Examples:
        >>> F = FormatDict
        >>> F1 = F({'font_name': 'Arial'})
        >>> F2 = F({'font_size': 12})
        >>> F3 = F({'font_name': 'Arial',  'font_size': 12})
        >>> F1 | F2 == F3
        True
        >>> F2 | F1 == F3
        True
        >>> hash(F1 | F2) == hash(F3)
        True
    """

    def __or__(self, other):
        return FormatDict({
            **self,
            **other
        })

    __ror__ = __or__

    def __hash__(self):
        return hash((*sorted(self.items()),))


@attrs(auto_attribs=True)
class FormatHandler(object):
    """This object is used to handle adding new formats when necessary. Only one should be used per Workbook."""
    target: XlsxWriterWorkbook
    _memoized: Dict[int, Format] = Factory(dict)

    def verify_format(self, format_: FormatDict):
        hashed = hash(format_)
        if hashed not in self._memoized:
            self._memoized[hashed] = self.target.add_format(format_)
        return self._memoized[hashed]


def ensure_format_uniqueness(class_):
    """A class decorator used to verify that all formats in the decorated class are unique and use FormatDict."""
    hashes = defaultdict(list)
    for attr in dir(class_):
        if not attr.startswith('_'):
            attr_value = getattr(class_, attr)
            if not isinstance(attr_value, FormatDict):
                raise TypeError(f'Format {attr_value} must be a FormatDict')
            hashes[hash(attr_value)].append(attr)

    for formats in hashes.values():
        if len(formats) > 1:
            raise ValueError(f'{formats} are the same')

    return class_


@ensure_format_uniqueness
class FormatsNamespace(object):
    base = FormatDict({})
    default_font_name = base | {'font_name': 'Liberation Sans'}
    default_font_size = base | {'font_size': 10}
    default_header_size = base | {'font_size': 18}

    percent = base | {'num_format': '0.0%'}
    regular_float = base | {'num_format': '0.00'}
    float_with_red = regular_float | {'num_format': '0.00;[RED]-0.00'}
    percent_with_red = percent | {'num_format': '0.0%;[RED]-0.0%'}

    left = base | {'align': 'left'}
    center = base | {'align': 'center', 'valign': 'vcenter'}
    right = base | {'align': 'right'}
    fill = base | {'align': 'fill'}
    justify = base | {'align': 'justify'}
    center_across = base | {'align': 'center_across'}
    distributed = base | {'align': 'distributed'}

    vbottom = base | {'valign': 'bottom'}
    vtop = base | {'valign': 'top'}
    vcenter = base | {'valign': 'vcenter'}
    vjustify = base | {'valign': 'vjustify'}
    vdistributed = base | {'valign': 'vdistributed'}

    rotated_0 = base | {'rotation': 0}
    rotated_90 = base | {'rotation': 90}
    rotated_180 = base | {'rotation': 270}
    rotated_270 = base | {'rotation': -90}

    wrapped = base | {'text_wrap': True}

    bold = base | {'bold': True}
    italic = base | {'italic': True}
    underline = base | {'underline': True}
    strikeout = base | {'font_strikeout': True}
    superscript = base | {'font_script': 1}
    subscript = base | {'font_script': 2}

    default_font = default_font_name | default_font_size | left
    default_font_bold = default_font | bold
    default_header = default_font_bold | default_header_size
    default_percent = default_font | percent | center
    default_percent_bold = default_percent | bold

    default_font_centered = default_font | center
    default_font_bold_centered = default_font_bold | center

    default_table_row_font = default_font_bold_centered | wrapped
    default_table_column_font = default_font_bold_centered | rotated_90 | vbottom

    left_border = base | {'left': 1}
    top_border = base | {'top': 1}
    right_border = base | {'right': 1}
    bottom_border = base | {'bottom': 1}
    highlight_border = left_border | top_border | right_border | bottom_border
