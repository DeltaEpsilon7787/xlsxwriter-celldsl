from functools import reduce
from itertools import chain
from typing import Iterable, List, Optional, TYPE_CHECKING

from attr import attrs
from xlsxwriter import Workbook
from xlsxwriter.worksheet import Worksheet

from .formats import FormatHandler

if TYPE_CHECKING:
    from cell_dsl import CommitTypes
    from ops.classes import WriteRichOp


@attrs(auto_attribs=True)
class WorkbookPair(object):
    """A pair used to bundle a :class:`FormatHandler` and a :ref:`Workbook <workbook>`"""
    wb: Workbook
    fmt: FormatHandler

    def add_worksheet(self, name):
        """Create a worksheet and bind it into a :class:`WorksheetTriplet`"""
        return WorksheetTriplet(self.wb, self.wb.add_worksheet(name), self.fmt)

    @classmethod
    def from_wb(cls, wb):
        """Bind a :class:`Workbook` into a :class:`WorkbookPair`"""
        return cls(wb, FormatHandler(wb))


@attrs(auto_attribs=True)
class WorksheetTriplet(object):
    """A triplet used for cell_dsl operations."""
    wb: Workbook
    ws: Worksheet
    fmt: FormatHandler


# noinspection PyPep8,PyPep8
def row_chain(
        iterable: Iterable,
        initial_save_name: Optional[str] = None,
        final_save_name: Optional[str] = None,
        range_name: Optional[str] = None,
        array_name: Optional[str] = None,
        step=1
) -> List['CommitTypes']:
    """Iterate `iterable` and insert `step` NextCols between each command and then come back to start
    Perhaps also save initial location as `initial_save_name` and final location as `final_save_name`.
    Perhaps also make the result a named range with name `range_name`.
    Perhaps add a forward reference as well with name `array_name`."""
    from . import ops

    return [
        ops.StackSave,
        initial_save_name is not None and ops.Save.at(initial_save_name) or None,
        [*chain.from_iterable([
            (command, *(ops.NextCol,) * step)
            for command in iterable
        ])][:-step],
        range_name is not None and (
            ops.DefineNamedRange
                .with_name(range_name)
                .top_left(-1)
        ) or None,
        array_name is not None and (
            ops.RefArray
                .at(array_name)
                .top_left(-1)
        ) or None,
        final_save_name is not None and ops.Save.at(final_save_name) or None,
        ops.StackLoad,
    ]


def col_chain(
        iterable: Iterable,
        initial_save_name: Optional[str] = None,
        final_save_name: Optional[str] = None,
        range_name: Optional[str] = None,
        array_name: Optional[str] = None,
        step=1
) -> List['CommitTypes']:
    """Iterate `iterable` and insert `step` NextRows between each command and then come back to start
    Perhaps also save initial location as `initial_save_name` and final location as `final_save_name`.
    Perhaps also make the result a named range with name `range_name`.
    Perhaps add a forward reference as well with name `array_name`."""
    from . import ops

    return [
        ops.StackSave,
        initial_save_name is not None and ops.Save.at(initial_save_name) or None,
        *[*chain.from_iterable([
            (command, *(ops.NextRow,) * step)
            for command in iterable
        ])][:-step],
        range_name is not None and (
            ops.DefineNamedRange
                .with_name(range_name)
                .top_left(-1)
        ) or None,
        array_name is not None and (
            ops.RefArray
                .at(array_name)
                .top_left(-1)
        ) or None,
        final_save_name is not None and ops.Save.at(final_save_name) or None,
        ops.StackLoad,
    ]


def chain_rich(iterable: Iterable['WriteRichOp']) -> 'WriteRichOp':
    """Take an `iterable` of `WriteRich` segments and combine them to produce a single WriteRich operation."""
    from .ops.classes import WriteRichOp

    return reduce(WriteRichOp.then, iterable)


def segment(name, iterable: Iterable['CommitTypes']) -> List['CommitTypes']:
    """A simple helper command to pad `iterable` of commands with SectionBegin and matching SectionEnd with `name`"""
    from .ops import SectionBegin, SectionEnd

    return [
        SectionBegin.with_name(name),
        iterable,
        SectionEnd
    ]
