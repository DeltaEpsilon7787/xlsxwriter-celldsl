import itertools
from typing import Any, Callable, ClassVar, Dict, List, Mapping, Optional, Set, Union

from attr import Factory, attrib, attrs, evolve
from xlsxwriter.utility import xl_range_formula

from . import traits
from .formats import FormatDict, FormatsNamespace
from .utils import WorksheetTriplet


class Command(object):
    """Base class for all commands."""
    pass


@attrs(auto_attribs=True, frozen=True)
class StackSaveOp(Command):
    """A command to push current location into save stack."""


@attrs(auto_attribs=True, frozen=True)
class StackLoadOp(Command):
    """A command to pop last location from save stack and jump to it."""


@attrs(auto_attribs=True, frozen=True)
class LoadOp(Command, traits.NamedPoint):
    """A command to jump to `at` a save point."""


@attrs(auto_attribs=True, frozen=True)
class SaveOp(Command, traits.NamedPoint):
    """A command to save current location `at` memory location."""


@attrs(auto_attribs=True, frozen=True)
class RefArrayOp(Command, traits.Range, traits.NamedPoint):
    """A forward reference to an array of cells with the name `name` defined using
    a rectangle with `top_left` and `bottom_right` specified. This is only used in charts.

    This is also a marker for such an array, meaning in commands that support forward references,
    you can use RefArray.at('name') to use a reference, which will be replaced
    with a string like '=SheetName!$C$1:$F$9'."""


@attrs(auto_attribs=True, frozen=True)
class SectionBeginOp(Command):
    """A command that does nothing, but may assist in debugging and documentation
    of scripts by giving providing segments in script `with_name`.

    During execution, if an error occurs, the surrounding names will be displayed, in order from most
    recent to least recent."""

    name: str = "__UNNAMED"

    def with_name(self, name: str):
        return evolve(self, name=name)


@attrs(auto_attribs=True, frozen=True)
class SectionEndOp(Command):
    """A command that indicates an end of the most recent `SectionBeginOp`."""


StackSave = StackSaveOp()
StackLoad = StackLoadOp()
Load = LoadOp()
Save = SaveOp()
RefArray = RefArrayOp()
SectionBegin = SectionBeginOp()
SectionEnd = SectionEndOp()


@attrs(auto_attribs=True, frozen=True)
class MoveOp(Command, traits.RelativePosition):
    """A command to move `r` rows and `c` columns away from current cell."""


@attrs(auto_attribs=True, frozen=True)
class AtCellOp(Command, traits.AbsolutePosition):
    """A command to go to a cell `r` and `c`."""


@attrs(auto_attribs=True, frozen=True)
class BacktrackCellOp(Command):
    """A command to `rewind` the position back in time. 0 stays in current cell, 1 goes to previous cell..."""
    n: int = 0

    def rewind(self, n_cells: int):
        return evolve(self, n=n_cells)


Move = MoveOp()
AtCell = AtCellOp()
BacktrackCell = BacktrackCellOp()


@attrs(auto_attribs=True, frozen=True)
class WriteOp(Command, traits.Data, traits.Format, traits.ExecutableCommand):
    """A command to write to this cell `with_data` and `with_format`."""

    def execute(self, target: WorksheetTriplet, coords: traits.Coords):
        return_code = target.ws.write(*coords, self.data, self.ensure_format(target.fmt))

        if return_code == -2:
            raise ValueError('Write failed because the string is longer than 32k characters')

        if return_code == -3:
            raise ValueError('Write failed because the URL is longer than 2079 characters long')

        if return_code == -4:
            raise ValueError('Write failed because there are more than 65530 URLs in the sheet')


@attrs(auto_attribs=True, frozen=True)
class MergeWriteOp(Command, traits.CardinalSize, traits.Data, traits.Format, traits.ExecutableCommand):
    """A command to merge `with_size` cols starting from current col and write `with_data` and `with_format`
    into this cell."""

    def execute(self, target: WorksheetTriplet, coords: traits.Coords):
        return_code = target.ws.merge_range(
            *coords,
            coords[0],
            coords[1] + self.size,
            self.data,
            self.ensure_format(target.fmt)
        )

        if return_code == -2:
            raise ValueError('Merge write failed because the string is longer than 32k characters')

        if return_code == -3:
            raise ValueError('Merge write failed because the URL is longer than 2079 characters long')

        if return_code == -4:
            raise ValueError('Merge write failed because there are more than 65530 URLs in the sheet')


@attrs(auto_attribs=True, frozen=True)
class WriteRichOp(Command, traits.Data, traits.Format, traits.ExecutableCommand):
    """A command to write a text run `with_data` and `with_format` to current position, `then` perhaps write some more,
    optionally `with_default_format`."""
    default_format: FormatDict = attrib(factory=FormatDict, converter=FormatDict)
    prev_fragment: Optional['WriteRichOp'] = None

    def execute(self, target: WorksheetTriplet, coords: traits.Coords):
        fragments: List[WriteRichOp] = self.rich_chain
        formats_and_data = itertools.chain.from_iterable(zip((
            fragment.ensure_format(target.fmt)
            for fragment in fragments
        ), (
            fragment.data
            for fragment in fragments
        )))

        return_code = target.ws.write_rich_string(*coords, *formats_and_data)
        if return_code == -2:
            raise ValueError('Rich write failed because the string is longer than 32k characters')

        if return_code == -4:
            raise ValueError('Rich write failed because of an empty string')

    def then(self, fragment: 'WriteRichOp'):
        """Submit additional fragments of the rich string"""
        if isinstance(fragment, WriteRichOp):
            return evolve(
                fragment,
                set_format=fragment.set_format or self.default_format,
                default_format=self.default_format,
                prev_fragment=self
            )
        else:
            raise TypeError

    def with_default_format(self, other):
        """Set format for fragments without a format. Should be applied to the first fragment"""
        return evolve(
            self,
            set_format=self.set_format or other,
            default_format=other
        )

    @property
    def rich_chain(self):
        """Not for public use; the flattened chain of segments"""
        chain = self

        result = []
        while chain.prev_fragment:
            result.append(chain)
            chain = chain.prev_fragment
        result.append(chain)
        result.reverse()

        return result

    @property
    def format_(self):
        """Not for public use; the format to be applied"""
        return self.set_format or self.default_format or self.FALLBACK_FORMAT


Write = WriteOp()
MergeWrite = MergeWriteOp()
WriteRich = WriteRichOp()


@attrs(auto_attribs=True, frozen=True)
class ImposeFormatOp(Command, traits.Format):
    """A command to append to merge current cell's format `with_format`."""
    set_format: FormatDict = attrib(default=FormatsNamespace.base, converter=FormatDict)


@attrs(auto_attribs=True, frozen=True)
class OverrideFormatOp(Command, traits.Format):
    """A command to override current cell's format `with_format`."""


ImposeFormat = ImposeFormatOp()
OverrideFormat = OverrideFormatOp()


@attrs(auto_attribs=True, frozen=True)
class DrawBoxBorderOp(Command, traits.Range):
    """Draw a box with borders where `top_left_point` and `bottom_right_point` are respective corners using
    `(right|top|left|bottom)_formats`."""

    right_format: Mapping = FormatsNamespace.right_border
    top_format: Mapping = FormatsNamespace.top_border
    left_format: Mapping = FormatsNamespace.left_border
    bottom_format: Mapping = FormatsNamespace.bottom_border

    def with_right_format(self, format_: Mapping):
        return evolve(self, right_format=format_)

    def with_top_format(self, format_: Mapping):
        return evolve(self, top_format=format_)

    def with_left_format(self, format_: Mapping):
        return evolve(self, left_format=format_)

    def with_bottom_format(self, format_: Mapping):
        return evolve(self, bottom_format=format_)


@attrs(auto_attribs=True, frozen=True)
class DefineNamedRangeOp(Command, traits.Range, traits.ExecutableCommand):
    """A command to make a box where `top_left_point` and `bottom_right_point` are respective corners of a range
    `with_name`."""

    name: str = "__DEFAULT"

    def with_name(self, name: str):
        return evolve(self, name=name)

    def execute(self, target: WorksheetTriplet, coords: traits.Coords):
        name = xl_range_formula(
            target.ws.name,
            *self.top_left_point,
            *self.bottom_right_point
        )

        target.wb.define_name(self.name, f'={name}')


DrawBoxBorder = DrawBoxBorderOp()
DefineNamedRange = DefineNamedRangeOp()


@attrs(auto_attribs=True, frozen=True)
class SetRowHeightOp(Command, traits.FractionalSize, traits.ExecutableCommand):
    """A command to set current row's height to `size`."""

    def execute(self, target: WorksheetTriplet, coords: traits.Coords):
        target.ws.set_row(row=coords[0], height=self.size)


@attrs(auto_attribs=True, frozen=True)
class SetColumnWidthOp(Command, traits.FractionalSize, traits.ExecutableCommand):
    """A command to set current column's height with `size`."""

    def execute(self, target: WorksheetTriplet, coords: traits.Coords):
        target.ws.set_column(coords[1], coords[1], self.size)


SetRowHeight = SetRowHeightOp()
SetColWidth = SetColumnWidthOp()


@attrs(auto_attribs=True, frozen=True)
class SubmitHPagebreakOp(Command):
    """A command to submit a horizontal page break at current row.
    This is preserved between several cell_dsl_context."""


@attrs(auto_attribs=True, frozen=True)
class SubmitVPagebreakOp(Command):
    """A command to submit a vertical page break at current row.
    This is preserved between several cell_dsl_context."""


@attrs(auto_attribs=True, frozen=True)
class ApplyPagebreaksOp(Command):
    """A command to apply all existing pagebreaks.
    Should come after all `SubmitHPagebreakOp` and `SubmitVPagebreakOp` have been committed."""


SubmitHPagebreak = SubmitHPagebreakOp()
SubmitVPagebreak = SubmitVPagebreakOp()
ApplyPagebreaks = ApplyPagebreaksOp()

NextRow = Move.r(1)
NextCol = Move.c(1)
PrevRow = Move.r(-1)
PrevCol = Move.c(-1)
NextRowSkip = Move.r(2)
NextColSkip = Move.c(2)
PrevRowSkip = Move.r(-2)
PrevColSkip = Move.c(-2)


@attrs(auto_attribs=True, frozen=True)
class AddCommentOp(Command, traits.Data, traits.ExecutableCommand):
    """A command to add a comment to this cell `with_data`, configured `with_options`"""
    options: Dict[str, Any] = Factory(dict)

    def with_options(self, options):
        return evolve(self, options={**self.options, **options})

    def execute(self, target: WorksheetTriplet, coords: traits.Coords):
        target.ws.write_comment(*coords, self.data, self.options)


AddComment = AddCommentOp()


@attrs(auto_attribs=True, frozen=True)
class AddChartOp(Command, traits.ExecutableCommand, traits.ForwardRef):
    """A command to add charts to the workbook, perhaps `with_subtype`,
    then `do` some method calls on the added chart."""
    type: str = 'bar'
    subtype: Optional[str] = None
    _sequence: List = Factory(list)

    _EXCEPTIONS: ClassVar[Set[str]] = {
        'set_style',
        'show_blanks_as',
        'show_hidden_data',
        'combine',
    }

    def do(self, func: Union[Callable, str], *args, **kwargs):
        """Add a deferred function call to the schedule.

        Args:
            func: Called function
                A method with the name of `func` will be called on the
                    chart object with `args` and `kwargs`
                Or this exact string name if provided.
            *args: Positional function call arguments
            **kwargs: Keyword function call arguments
        """
        func_name = func

        if callable(func):
            func_name = func.__name__

        self._sequence.append((func_name, args, kwargs))

        return self

    def nop_do(self, func: Union[Callable, str], *args, **kwargs):
        """Same as `do`, but returns NoOp"""

        self.do(func, *args, **kwargs)

    def with_subtype(self, subtype):
        return evolve(self, subtype=subtype)

    def execute(self, target: WorksheetTriplet, coords: traits.Coords):
        result = target.wb.add_chart({
            'type': self.type,
            'subtype': self.subtype
        })

        def ref_expander(options: Mapping[str, Any]):
            def recursive(current_node: Mapping[str, Any]):
                for key, value in current_node.items():
                    if isinstance(value, RefArrayOp):
                        string_repr = self.resolved_refs[value.point_name]
                        string_repr = f"'{target.ws.name}'!{string_repr}"
                        yield key, string_repr
                    elif isinstance(value, Mapping):
                        yield key, dict(recursive(value))
                    else:
                        yield key, value

            return dict(recursive(options))

        for (f, a, k) in self._sequence:
            target_func = getattr(result, f)
            if target_func is None:
                raise ValueError(f'Object of type {type(result)} does not have a method called {f}')

            if f not in self._EXCEPTIONS:
                if k.get('options') is not None:
                    options = k['options']
                    k = {
                        key: value
                        for key, value in k.items()
                        if key != 'options'
                    }
                else:
                    options = a[0]

                new_options = ref_expander(options)

                target_func(new_options, *a[1:], **k)
            else:
                target_func(*a, **k)

        target.ws.insert_chart(*coords, result)


AddAreaChart = AddChartOp(type='area')
AddBarChart = AddChartOp(type='bar')
AddColumnChart = AddChartOp(type='column')
AddLineChart = AddChartOp(type='line')
AddPieChart = AddChartOp(type='pie')
AddDoughnutChart = AddChartOp(type='doughnut')
AddScatterChart = AddChartOp(type='scatter')
AddStockChart = AddChartOp(type='stock')
AddRadarChart = AddChartOp(type='radar')
