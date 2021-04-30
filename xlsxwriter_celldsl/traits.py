from abc import abstractmethod
from functools import lru_cache
from numbers import Integral, Real
from typing import Any, ClassVar, Dict, Optional, TYPE_CHECKING, Tuple, TypeVar, Union

from attr import Factory, attrib, attrs, evolve

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


@attrs(auto_attribs=True)
class FractionalSize(Trait):
    """Commands with this trait require some kind of real-valued `size`"""
    size: Real = 0.0

    def with_size(self: T, size: Real) -> T:
        return evolve(self, size=size)


@attrs(auto_attribs=True)
class CardinalSize(Trait):
    """Commands with this trait require some kind of integer `size`"""
    size: int = 0

    def with_size(self: T, size: Integral) -> T:
        return evolve(self, size=size)


@attrs(auto_attribs=True)
class Position(Trait):
    """Commands with this aspect target a specific cell at `row` and `col`"""

    row: int = -1
    col: int = -1

    @lru_cache(maxsize=None)
    def at_row(self: T, row: Integral) -> T:
        return evolve(self, row=row)

    @lru_cache(maxsize=None)
    def at_col(self: T, col: Integral) -> T:
        return evolve(self, col=col)


@attrs(auto_attribs=True)
class Range(Trait):
    """Commands with this trait target a range of cells.

    Attributes:
        top_left_point: Union[str, Integral, Coords]
            Top left corner of the bounding box.
        bottom_right_point: Union[str, Integral, Coords]
            Bottom right corner of the bounding box.

    Notes:
        For both `top_left_point` and `bottom_right_point`, the types
        define how to locate the target cell:
            If it's a string, this will be the save point name at which the save occurs.
            If it's an Integral
                If positive: this will be last n-th visited cell.
                If negative: this will be last n-th position in save stack which
                    will be retrieved without popping it from the stack.
                If zero, current cell will be the target.
            If it's Coords, it will be the absolute coords of the cell.
        """
    top_left_point: CellPointer = 0
    bottom_right_point: CellPointer = 0

    def top_left(self, point: CellPointer):
        return evolve(self, top_left_point=point)

    def bottom_right(self, point: CellPointer):
        return evolve(self, bottom_right_point=point)


@attrs(auto_attribs=True)
class RelativeMove(Trait):
    """Commands with this trait target a cell relative to the current cell."""
    delta_row: Integral = 0
    delta_col: Integral = 0

    @lru_cache(maxsize=None)
    def r(self: T, delta_row: Integral) -> T:
        return evolve(self, delta_row=delta_row)

    @lru_cache(maxsize=None)
    def c(self: T, delta_col: Integral) -> T:
        return evolve(self, delta_col=delta_col)


@attrs(auto_attribs=True)
class Data(Trait):
    """Commands with this trait provide some `data` to the function for writing."""
    data: Any = ""

    def with_data(self: T, data: Any) -> T:
        return evolve(self, data=data)


@attrs(auto_attribs=True)
class Format(Trait):
    """Commands with this trait provide some `format` to the function for writing."""
    FALLBACK_FORMAT: ClassVar[FormatDict] = FormatsNamespace.default_font
    set_format: Optional[FormatDict] = attrib(default=None)

    def with_format(self: T, format_) -> T:
        return evolve(self, set_format=self.format_ | format_)

    def ensure_format(self, handler: FormatHandler):
        return handler.verify_format(self.format_)

    def set_default_font(self: T, format_):
        self.__class__.FALLBACK_FORMAT = FormatDict(format_)

    @property
    def format_(self) -> FormatDict:
        return self.set_format or self.FALLBACK_FORMAT


@attrs(auto_attribs=True)
class NamedPoint(Trait):
    """Commands with this trait give a temporary name to the cell to be referenced later."""
    point_name: str = "__DEFAULT"

    def at(self: T, point_name) -> T:
        return evolve(self, point_name=point_name)


@attrs(auto_attribs=True)
class ForwardRef(Trait):
    """Commands with this trait manage a list of forward array references."""
    resolved_refs: Dict[str, str] = Factory(dict)

    def inject_refs(self: T, ref_array) -> T:
        return evolve(self, resolved_refs=ref_array)


class ExecutableCommand(Trait):
    """Commands with this trait are active commands representing actions that will be done in the worksheet."""

    @abstractmethod
    def execute(self, target: WorksheetTriplet, coords: Coords):
        raise NotImplemented
