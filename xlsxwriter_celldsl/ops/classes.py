import itertools
from io import BytesIO
from pathlib import Path
from typing import Any, Dict, Generic, List, Mapping, Optional, Tuple, TypeVar, Union

from attr import Factory, attrib, attrs, evolve
from xlsxwriter.utility import xl_range_formula

from formats import FormatDict, FormatsNamespace
from ops import traits
from utils import WorksheetTriplet

T = TypeVar('T')


class Command(object):
    """Base class for all commands."""
    OVERWRITE_SENSITIVE = False
    pass


@attrs(auto_attribs=True, frozen=True, order=False)
class StackSaveOp(Command):
    """A command to push current location into save stack."""


@attrs(auto_attribs=True, frozen=True, order=False)
class StackLoadOp(Command):
    """A command to pop last location from save stack and jump to it."""


@attrs(auto_attribs=True, frozen=True, order=False)
class LoadOp(Command, traits.NamedPoint):
    """A command to jump to :func:`at` a save point."""


@attrs(auto_attribs=True, frozen=True, order=False)
class SaveOp(Command, traits.NamedPoint):
    """A command to save current location :func:`at` memory location."""


@attrs(auto_attribs=True, frozen=True, order=False)
class RefArrayOp(Command, traits.Range, traits.NamedPoint):
    """A forward reference to an array of cells :func:`with_name` defined using
    a rectangle with :func:`top_left` and :func:`bottom_right` specified. This is only used in charts.

    This is also a marker for such an array, meaning in commands that support forward references,
    you can use RefArray.at('name') to use a reference, which will be replaced
    with a string like ``'=SheetName!$C$1:$F$9'``."""


@attrs(auto_attribs=True, frozen=True, order=False)
class SectionBeginOp(Command):
    """A command that does nothing, but may assist in debugging and documentation
    of scripts by giving providing segments in script :func:`with_name`.

    During execution, if an error occurs, the surrounding names will be displayed, in order from most
    recent to least recent."""

    name: str = "__UNNAMED"

    def with_name(self, name: str):
        return evolve(self, name=name)


@attrs(auto_attribs=True, frozen=True, order=False)
class SectionEndOp(Command):
    """A command that indicates an end of the most recent `SectionBeginOp`."""


@attrs(auto_attribs=True, frozen=True, order=False)
class MoveOp(Command, traits.RelativePosition):
    """A command to move :func:`r` rows and :func:`c` columns away from current cell."""


@attrs(auto_attribs=True, frozen=True, order=False)
class AtCellOp(Command, traits.AbsolutePosition):
    """A command to go to a cell :func:`r` and :func:`c`."""


@attrs(auto_attribs=True, frozen=True, order=False)
class BacktrackCellOp(Command):
    """A command to :func:`rewind` the position back in time. 0 stays in current cell, 1 goes to previous cell..."""
    n: int = 0

    def rewind(self, n_cells: int):
        return evolve(self, n=n_cells)


@attrs(auto_attribs=True, frozen=True, order=False)
class WriteOp(Command, traits.Data, traits.DataType, traits.Format, traits.ExecutableCommand):
    """A command to write to this cell :func:`with_data` with data :func:`with_data_type` and :func:`with_format`."""
    OVERWRITE_SENSITIVE = True

    def execute(self, target: WorksheetTriplet, coords: traits.Coords):
        args = *coords, self.data, self.ensure_format(target.fmt)
        if self.data_type is not None:
            return_code = getattr(target.ws, f'write_{self.data_type}')(*args)
        else:
            return_code = target.ws.write(*args)

        if return_code == -2:
            raise ValueError('Write failed because the string is longer than 32k characters')

        if return_code == -3:
            raise ValueError('Write failed because the URL is longer than 2079 characters long')

        if return_code == -4:
            raise ValueError('Write failed because there are more than 65530 URLs in the sheet')


@attrs(auto_attribs=True, frozen=True, order=False)
class MergeWriteOp(Command, traits.CardinalSize, traits.Data, traits.DataType, traits.Format, traits.ExecutableCommand):
    """
    A command to merge :func:`with_size` cols starting from current col and
    write :func:`with_data` with data :func:`with_data_type` and :func:`with_format` into this cell.
    """

    OVERWRITE_SENSITIVE = True

    def execute(self, target: WorksheetTriplet, coords: traits.Coords):
        return_code = target.ws.merge_range(
            *coords,
            coords[0],
            coords[1] + self.size,
            self.data,
            self.ensure_format(target.fmt)
        )

        if self.data_type is not None:
            # In order to force a data type into a merged cell, we have to perform a second write
            #   as shown here: <https://xlsxwriter.readthedocs.io/example_merge_rich.html>
            return_code = getattr(target.ws, f"write_{self.data_type}")(
                *coords,
                self.data,
                self.ensure_format(target.fmt)
            )

        if return_code == -2:
            raise ValueError('Merge write failed because the string is longer than 32k characters')

        if return_code == -3:
            raise ValueError('Merge write failed because the URL is longer than 2079 characters long')

        if return_code == -4:
            raise ValueError('Merge write failed because there are more than 65530 URLs in the sheet')


@attrs(auto_attribs=True, frozen=True, order=False)
class WriteRichOp(Command, traits.Data, traits.Format, traits.ExecutableCommand):
    """A command to write a text run :func:`with_data` and
    :func:`with_format` to current position, :func:`then` perhaps write some more,
    optionally :func:`with_default_format`."""
    default_format: FormatDict = attrib(factory=FormatDict, converter=FormatDict)
    prev_fragment: Optional['WriteRichOp'] = None
    OVERWRITE_SENSITIVE = True

    def execute(self, target: WorksheetTriplet, coords: traits.Coords):
        fragments: List[WriteRichOp] = self.rich_chain
        formats_and_data = [*itertools.chain.from_iterable(zip((
            fragment.ensure_format(target.fmt)
            for fragment in fragments
        ), (
            fragment.data
            for fragment in fragments
        )))]

        return_code = target.ws.write_rich_string(*coords, *formats_and_data)
        if return_code == -5:
            return_code = target.ws.write_string(*coords, formats_and_data[1], formats_and_data[0])

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
            raise TypeError(fragment)

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


@attrs(auto_attribs=True, frozen=True, order=False)
class ImposeFormatOp(Command, traits.Format):
    """A command to append to merge current cell's format :func:`with_format`."""
    set_format: FormatDict = attrib(default=FormatsNamespace.base, converter=FormatDict)


@attrs(auto_attribs=True, frozen=True, order=False)
class OverrideFormatOp(Command, traits.Format):
    """A command to override current cell's format :func:`with_format`."""


@attrs(auto_attribs=True, frozen=True, order=False)
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


@attrs(auto_attribs=True, frozen=True, order=False)
class DefineNamedRangeOp(Command, traits.Range, traits.ExecutableCommand):
    """A command to make a box where `top_left_point` and `bottom_right_point` are respective corners of a range
    :func:`with_name`."""

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


@attrs(auto_attribs=True, frozen=True, order=False)
class SetRowHeightOp(Command, traits.FractionalSize, traits.ExecutableCommand):
    """A command to set current row's height :func:`with_size`."""

    def execute(self, target: WorksheetTriplet, coords: traits.Coords):
        target.ws.set_row(row=coords[0], height=self.size)


@attrs(auto_attribs=True, frozen=True, order=False)
class SetColumnWidthOp(Command, traits.FractionalSize, traits.ExecutableCommand):
    """A command to set current column's height with :func:`with_size`."""

    def execute(self, target: WorksheetTriplet, coords: traits.Coords):
        target.ws.set_column(coords[1], coords[1], self.size)


@attrs(auto_attribs=True, frozen=True, order=False)
class SubmitHPagebreakOp(Command):
    """A command to submit a horizontal page break at current row.
    This is preserved between several cell_dsl_context."""


@attrs(auto_attribs=True, frozen=True, order=False)
class SubmitVPagebreakOp(Command):
    """A command to submit a vertical page break at current row.
    This is preserved between several cell_dsl_context."""


@attrs(auto_attribs=True, frozen=True, order=False)
class ApplyPagebreaksOp(Command):
    """A command to apply all existing pagebreaks.
    Should come after all `SubmitHPagebreakOp` and `SubmitVPagebreakOp` have been committed."""


@attrs(auto_attribs=True, frozen=True, order=False)
class AddCommentOp(Command, traits.Data, traits.Options, traits.ExecutableCommand):
    """A command to add a comment to this cell :func:`with_data`, configured :func:`with_options`."""

    def execute(self, target: WorksheetTriplet, coords: traits.Coords):
        target.ws.write_comment(*coords, self.data, self.options)


class _ChartHelper:
    def __getattr__(self, item):
        def recorder(*args, **kwargs):
            return item, args, kwargs

        return recorder


_CHART_FUNC_EXCEPTIONS = {
    'set_style',
    'show_blanks_as',
    'show_hidden_data',
    'combine',
}


@attrs(auto_attribs=True, frozen=True, order=False)
class AddChartOp(Command, traits.ExecutableCommand, traits.ForwardRef, Generic[T]):
    """
    A command to add a chart to this cell perhaps :func:`with_subtype` and then :func:`do`
    call some methods on the associated `target` class.
    """
    type: str = 'bar'
    subtype: Optional[str] = None

    # It's typed T to trick PyCharm or other similar systems
    #   to autocomplete using methods of the associated
    #   class. It is also appropriate since
    #   ChartHelper works kinda like a mock of the associated
    #   class.
    target: T = Factory(_ChartHelper)
    action_chain: List[Tuple[str, Tuple[Any, ...], Dict[str, Any]]] = Factory(list)

    def with_subtype(self, subtype):
        return evolve(self, subtype=subtype)

    def do(self, command_list):
        """
        Add `command_list` to `action_chain`

        Example:
            >>> from xlsxwriter_celldsl.ops import AddLineChart, AddBarChart, RefArray
            ... AddLineChart.do([
            ...     # You really should only use `target` attribute of this class
            ...     AddLineChart.target.add_series({'values': '=SheetName!$A$1:$D$1'}),
            ...     # Charts allow to use `RefArray` in place of literal cell ranges
            ...     AddLineChart.target.add_series({'values': RefArray.at('some ref')}),
            ...     # Combine method accepts AddChartOp
            ...     AddLineChart.target.combine(
            ...         # This will combine this line chart with a bar chart
            ...         AddBarChart.do([])
            ...     )
            ... ])
        """
        return evolve(self, action_chain=[*self.action_chain, *command_list])

    def execute(self, target: WorksheetTriplet, coords: traits.Coords, _is_secondary=False):
        result = target.wb.add_chart({
            'type': self.type,
            'subtype': self.subtype
        })

        def ref_expander(source: Mapping[str, Any]):
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

            return dict(recursive(source))

        for (f, a, k) in self.action_chain:
            target_func = getattr(result, f)

            if f not in _CHART_FUNC_EXCEPTIONS:
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
            elif f == 'combine':
                # Special case, need to create another chart
                secondary, = a
                secondary = secondary.execute(target, coords, _is_secondary=True)

                result.combine(secondary)
            else:
                target_func(*a, **k)

        if _is_secondary:
            return result
        else:
            target.ws.insert_chart(*coords, result)


@attrs(auto_attribs=True, frozen=True, order=False)
class AddConditionalFormatOp(Command, traits.ExecutableCommand, traits.Range, traits.Options, traits.Format):
    """A command to add a conditional format to a range of cells with :func:`top_left` and :func:`bottom_right`
    corners parametrized :func:`with_options`.

    To configure the format, you can either use :func:`with_format` or specify the :class:`FormatDict` as
    format key in options. Do not use both at the same time however."""

    def execute(self, target: WorksheetTriplet, coords: traits.Coords):
        if self.set_format and self.options.get('format'):
            raise ValueError('Both format key and format field are specified, use only one of them.')

        fmt = None
        if self.set_format:
            fmt = self.ensure_format(target.fmt)
        if self.options.get('format'):
            fmt = self.with_format(self.options['format']).ensure_format(target.fmt)

        result = target.ws.conditional_format(
            *self.top_left_point,
            *self.bottom_right_point,
            {
                **self.options,
                **(fmt and {'format': fmt} or {})
            }
        )

        if result == -2:
            raise ValueError(f"Invalid parameter or options: {self.options}")


@attrs(auto_attribs=True, frozen=True, order=False)
class AddImageOp(Command, traits.ExecutableCommand, traits.Options):
    """A command to add an image to this cell either getting it :func:`with_filepath` or
    constructing it :func:`with_image_data`, parametrized :func:`with_options`."""
    file_path: str = ""

    def with_filepath(self, file_path: Union[Path, str]):
        return evolve(self, file_path=str(file_path))

    def with_image_data(self, image_data: BytesIO):
        return self.with_options({'image_data': image_data})

    def execute(self, target: WorksheetTriplet, coords: traits.Coords):
        target.ws.insert_image(*coords, self.file_path, self.options)
