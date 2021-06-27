from abc import abstractmethod
from numbers import Integral, Real
from typing import Any, ClassVar, Dict, Optional, TYPE_CHECKING, Tuple, TypeVar, Union

from attr import attrib, attrs, evolve

from .formats import FormatDict, FormatHandler, FormatsNamespace
from .utils import WorksheetTriplet

if TYPE_CHECKING:
    pass

T = TypeVar('T')

Coords = Tuple[int, int]
CellPointer = Union[str, int, Coords]


@attrs(auto_attribs=True)
class Trait(object):
    pass


@attrs(auto_attribs=True, frozen=True, order=False)
class FractionalSize(Trait):
    """Commands with this trait require some kind of real-valued `size`"""
    size: float = 0.0

    def with_size(self: T, size: Real) -> T:
        """Specify the real-valued `size` for this object."""
        return evolve(self, size=float(size))


@attrs(auto_attribs=True, frozen=True, order=False)
class CardinalSize(Trait):
    """Commands with this trait require some kind of integer `size`"""
    size: int = 0

    def with_size(self: T, size: Integral) -> T:
        """Specify the integral `size` for this object."""
        return evolve(self, size=int(size))


@attrs(auto_attribs=True, frozen=True, order=False)
class AbsolutePosition(Trait):
    """Commands with this aspect target a specific cell at `row` and `col`"""

    row: int = -1
    col: int = -1

    def r(self: T, row: Integral) -> T:
        """Target the cell at `row` for this object."""
        return evolve(self, row=int(row))

    def c(self: T, col: Integral) -> T:
        """Target the cell at `col` for this object."""
        return evolve(self, col=int(col))


@attrs(auto_attribs=True, frozen=True, order=False)
class RelativePosition(Trait):
    """Commands with this trait target a cell relative to the current cell."""
    row: int = 0
    col: int = 0

    def r(self: T, delta_row: Integral) -> T:
        """Specify a target `delta_row` away from current position for this object."""
        return evolve(self, row=int(delta_row))

    def c(self: T, delta_col: Integral) -> T:
        """Specify a target `delta_col` away from current position for this object."""
        return evolve(self, col=int(delta_col))


@attrs(auto_attribs=True, frozen=True, order=False)
class Range(Trait):
    """
    Commands with this trait target a range of cells bounded by a box with `top_left_point` and `bottom_right_point`.

    Notes:
        `top_left_point` and `bottom_right_point` have different behavior depending on the value type

        * If it's a string, this will be the save point name at which the save occurs.
        * If it's an Integral

          * If positive: this will be last n-th visited cell.

          * If negative: this will be last n-th position in save stack which will be retrieved without popping it
            from the stack.

          * If zero, current cell will be the target.

        * If it's Coords, it will be the absolute coords of the cell.
    """
    top_left_point: CellPointer = 0
    bottom_right_point: CellPointer = 0

    def top_left(self, point: CellPointer):
        """Specify top left corner `point` for this object."""
        return evolve(self, top_left_point=point)

    def bottom_right(self, point: CellPointer):
        """Specify bottom right corner `point` for this object."""
        return evolve(self, bottom_right_point=point)


@attrs(auto_attribs=True, frozen=True, order=False)
class Data(Trait):
    """Commands with this trait provide some `data` to the function for writing."""
    data: Any = ""

    def with_data(self: T, data: Any) -> T:
        """Specify the `data` for this object."""
        return evolve(self, data=data)


@attrs(auto_attribs=True, frozen=True, order=False)
class DataType(Trait):
    """Commands with this trait parametrize `data_type`, generally assuming :class:`Data` trait is also present."""
    data_type: Optional[str] = None
    ACCEPTED_DATA_TYPES = {None, 'string', 'number', 'blank', 'formula', 'datetime', 'boolean', 'url'}

    def with_data_type(self, data_type: Optional[str]):
        """Force the write to use a specific method or default to generic `write` if `data_type` is None.

        Accepted values are: string, number, blank, formula, datetime, boolean, url and None

        See Also:
            :func:`write`
        """
        if data_type not in self.ACCEPTED_DATA_TYPES:
            raise ValueError(
                f'Data type {data_type} is not valid: valid values are {self.ACCEPTED_DATA_TYPES}'
            )

        return evolve(self, data_type=data_type)


@attrs(auto_attribs=True, frozen=True, order=False)
class Format(Trait):
    """Commands with this trait provide some `format_` to the function for writing."""
    FALLBACK_FORMAT: ClassVar[FormatDict] = FormatsNamespace.default_font
    set_format: Optional[FormatDict] = attrib(default=None, repr=False)

    def with_format(self: T, format_) -> T:
        """Specify the `format_` for this object."""
        return evolve(self, set_format=self.format_ | format_)

    def ensure_format(self, handler: FormatHandler):
        """Not for public use; inject this format into the workbook using format `handler`."""
        return handler.verify_format(self.format_)

    def set_base_format(self: T, format_):
        """
        Set the fallback format globally for the entire project.
        This format will be used in absence of `set_format` and all formats later will merge with it."""
        self.__class__.FALLBACK_FORMAT = FormatDict(format_)

    @property
    def format_(self) -> FormatDict:
        return self.set_format or self.FALLBACK_FORMAT


@attrs(auto_attribs=True, frozen=True, order=False)
class NamedPoint(Trait):
    """Commands with this trait give a temporary `point_name` to the cell to be referenced later."""
    point_name: str = "__DEFAULT"

    def at(self: T, point_name) -> T:
        """Specify the `point_name` for this object."""
        return evolve(self, point_name=point_name)


@attrs(auto_attribs=True, frozen=True, order=False)
class ForwardRef(Trait):
    """Commands with this trait manage a list of `resolve_refs`."""
    resolved_refs: Dict[str, str] = attrib(factory=dict, repr=False, hash=False)

    def inject_refs(self: T, ref_array) -> T:
        """Not for public use; inject the `ref_array` into this object."""
        return evolve(self, resolved_refs=ref_array)


@attrs(auto_attribs=True, frozen=True, order=False)
class Options(Trait):
    """Commands with this trait have a dictionary of `options`"""

    options: Dict[str, Any] = attrib(factory=dict, repr=False, hash=False)

    def with_options(self, options: Dict[str, Any]):
        """Add new `options` to our configuration."""
        return evolve(self, options={**self.options, **options})


class ExecutableCommand(Trait):
    """Commands with this trait are active commands representing actions that will be done in the worksheet."""

    @abstractmethod
    def execute(self, target: WorksheetTriplet, coords: Coords):
        """Not for public use; execute the command at `coords` in `target`."""
        raise NotImplemented
