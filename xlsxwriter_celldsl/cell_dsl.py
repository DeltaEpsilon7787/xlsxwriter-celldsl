from collections import defaultdict, deque
from contextlib import contextmanager
from itertools import groupby
from operator import itemgetter
from pprint import pformat
from typing import ContextManager, DefaultDict, Dict, Iterable, Iterator, List, Optional, Tuple, Union, cast
from warnings import warn

from attr import Factory, attrib, attrs, evolve
from xlsxwriter.utility import xl_range_abs

from xlsxwriter_celldsl import ops
from xlsxwriter_celldsl.formats import FormatDict, FormatsNamespace as F
from xlsxwriter_celldsl.ops import Move, Write, WriteRich
from xlsxwriter_celldsl.traits import Coords, ExecutableCommand, ForwardRef, Range
from xlsxwriter_celldsl.utils import WorksheetTriplet

MovementShortForm = int
WriteDataShortForm = str
FormatAsDictForm = dict

CommitTypes = Union[
    Iterable['CommitTypes'],
    ops.Command,
    MovementShortForm,
    WriteDataShortForm,
    FormatAsDictForm,
    FormatDict,
    None
]


def name_stack_repr(name_stack):
    segments = []
    for name, group in groupby(name_stack):
        group_len = len([*group])
        segments.append(group_len > 1 and f"{name}x{group_len}" or name)
    return segments[::-1]


class CellDSLError(Exception):
    """Base Cell DSL error"""

    def __init__(self, message, action_num=None, action_lst=None, name_stack=None, action=None, save_points=None):
        self.message = message
        self.action_num = action_num
        self.action_lst = action_lst
        self.name_stack = name_stack
        self.action = action
        self.save_points = save_points

    def __str__(self):
        segments = []
        if self.name_stack is not None:
            segments.append(f"Name stack: {pformat(name_stack_repr(self.name_stack))}")
        if self.action_lst is not None and self.action_num is not None:
            segments.append(
                f"Adjacent actions: {pformat(self.action_lst[max(self.action_num - 10, 0):self.action_num + 10])}"
            )
        if self.action_num is not None:
            segments.append(f"Action num: {pformat(self.action_num)}")
        if self.action is not None:
            segments.append(f"Triggering action: {self.action}")
        if self.save_points is not None:
            segments.append(f"Save points already present: {pformat(self.save_points)}")
        additional_info = "\n".join(segments)

        full_message = [self.message]
        if additional_info:
            full_message.append(f"Additional info:\n{additional_info}")

        return "\n".join(full_message)


class MovementCellDSLError(CellDSLError):
    """A Cell DSL error triggered by invalid movement"""


class ExecutionCellDSLError(Exception):
    """A Cell DSL error triggered by an exception during execution."""


@attrs(auto_attribs=True)
class StatReceiver(object):
    """Receiver object for stats after cell_dsl_context context finishes execution.

    Attributes:
        initial_row: Row where the execution started
        initial_col: Column where the execution started
        coord_pairs:
            A list of pairs, with the first element being the coordinate pair and
            the second being the command executed.
        save_points:
            A dictionary where the key is the name of the save point and the value is
            the coordinates of the cell the save point points to.
    """
    initial_row: int = attrib(init=False)
    initial_col: int = attrib(init=False)
    coord_pairs: List[Tuple[Coords, ops.Command]] = attrib(init=False)
    save_points: Dict[str, Coords] = attrib(init=False, factory=dict)

    @property
    def is_null(self):
        return len(self.coord_pairs) < 2

    @property
    def _coord_iter(self) -> Iterator[Coords]:
        return cast(Iterator[Coords], map(itemgetter(0), self.coord_pairs))

    @property
    def max_row(self) -> int:
        return cast(int, max(self._coord_iter, key=itemgetter(0)))

    @property
    def max_col(self) -> int:
        return cast(int, max(self._coord_iter, key=itemgetter(1)))

    @property
    def max_row_at_max_col(self):
        max_col = self.max_col

        return max(
            row
            for row, col in self._coord_iter
            if col == max_col
        )

    @property
    def max_col_at_max_row(self):
        max_row = self.max_row

        return max(
            col
            for row, col in self._coord_iter
            if row == max_row
        )

    @property
    def max_coords(self) -> Coords:
        return self.max_row, self.max_col

    @property
    def initial_coords(self) -> Coords:
        return self.initial_row, self.initial_col


def _process_movement(action_list, row, col) -> Tuple[DefaultDict[Coords, List[ops.Command]], Dict[str, Coords]]:
    result = defaultdict(list)
    save_points = {}
    save_stack = deque()
    name_stack = deque()

    visited = [(row, col)]

    def trigger_movement_error(message, exc=None):
        raise MovementCellDSLError(
            message,
            action_num,
            action_list,
            name_stack,
            action,
            save_points
        ) from exc

    for action_num, action in enumerate(action_list):
        action_type = type(action)
        if action_type is ops.LoadOp:
            try:
                row, col = save_points[action.point_name]
            except KeyError as e:
                trigger_movement_error(f'Save point {action.point_name} does not exist.', e)
            visited.append((row, col))
        elif action_type is ops.StackLoadOp:
            try:
                row, col = save_stack.pop()
            except IndexError as e:
                trigger_movement_error(f'Save stack is empty.', e)
            visited.append((row, col))
        elif action_type is ops.SaveOp:
            save_points[action.point_name] = (row, col)
        elif action_type is ops.StackSaveOp:
            save_stack.append((row, col))
        elif action_type is ops.MoveOp:
            row += action.row
            col += action.col
            visited.append((row, col))
        elif action_type is ops.AtCellOp:
            row = action.row
            col = action.col
            visited.append((row, col))
        elif action_type is ops.BacktrackCellOp:
            try:
                for _ in range(action.n + 1):
                    row, col = visited.pop()
            except IndexError as e:
                trigger_movement_error(f'Could not backtrack {action.n} cells.', e)
        else:
            if isinstance(action, Range):
                if isinstance(action.top_left_point, int):
                    if action.top_left_point > 0:
                        try:
                            action = action.top_left(visited[-action.top_left_point - 1])
                        except IndexError as e:
                            trigger_movement_error(
                                f'Top left corner would use {action.top_left_point} last visited cell, but only '
                                f'{len(visited)} cells have been visited',
                                e
                            )
                    elif action.top_left_point < 0:
                        try:
                            action = action.top_left(save_stack[action.top_left_point])
                        except IndexError as e:
                            trigger_movement_error(
                                f'Top left corner would look {-action.top_left_point} positions'
                                f'up the save stack, but there is only '
                                f'{len(save_stack)} saves',
                                e
                            )
                    else:
                        action = action.top_left((row, col))
                if isinstance(action.bottom_right_point, int):
                    if action.bottom_right_point > 0:
                        try:
                            action = action.bottom_right(visited[-action.bottom_right_point - 1])
                        except IndexError as e:
                            trigger_movement_error(
                                f'Bottom right corner would use {action.bottom_right_point} last visited cell, but only '
                                f'{len(visited)} cells have been visited',
                                e
                            )
                    elif action.bottom_right_point < 0:
                        try:
                            action = action.bottom_right(save_stack[action.bottom_right_point])
                        except IndexError as e:
                            trigger_movement_error(
                                f'Bottom right corner would look {-action.bottom_right_point} positions'
                                f'up the save stack, but there is only '
                                f'{len(save_stack)} saves',
                                e
                            )
                    else:
                        action = action.bottom_right((row, col))
            elif action_type is ops.SectionBeginOp:
                name_stack.append(action.name)
            elif action_type is ops.SectionEndOp:
                name_stack.pop()

            result[(row, col)].append(action)

        # 2^20 and 2^14 are Excel limits for the amount of row and columns respectively.
        if row not in range(0, 2 ** 20) or col not in range(0, 2 ** 14):
            trigger_movement_error(f'Illegal coords have been reached, this is not allowed.')

    if len(name_stack) > 0:
        raise trigger_movement_error(f'Name stack is not empty, every SectionBegin must be matched with SectionEnd')

    return result, save_points


def _inject_coords(coord_action_map, save_points) -> Tuple[DefaultDict[Coords, List[ops.Command]], Dict[str, str]]:
    result = defaultdict(list)
    ref_array = {}

    def trigger_cell_dsl_error(message, exc=None):
        raise CellDSLError(message, action=action, save_points=save_points) from exc

    for coords, actions in coord_action_map.items():
        for action in actions:
            if isinstance(action, Range):
                if isinstance(action.top_left_point, str):
                    try:
                        action = action.top_left(save_points[action.top_left_point])
                    except KeyError as e:
                        trigger_cell_dsl_error(
                            f"Tried to use a save point named {action.top_left_point} "
                            f"for top left corner, but it doesn't exist",
                            e
                        )
                if isinstance(action.bottom_right_point, str):
                    try:
                        action = action.bottom_right(save_points[action.bottom_right_point])
                    except KeyError as e:
                        trigger_cell_dsl_error(
                            f"Tried to use a save point named {action.bottom_right_point} "
                            f"for bottom right corner, but it doesn't exist",
                            e
                        )

            if isinstance(action, ops.RefArrayOp):
                c = xl_range_abs(*action.top_left_point, *action.bottom_right_point)
                ref_array[action.point_name] = f'{c}'
                continue

            result[coords].append(action)

    return result, ref_array


def _introduce_ref_arrays(coord_action_map, ref_array):
    result: DefaultDict[Coords, List[ops.Command]] = defaultdict(list)

    for coords, actions in coord_action_map.items():
        for action in actions:
            if isinstance(action, ForwardRef):
                action = action.inject_refs(ref_array)

            result[coords].append(action)

    return result


def _expand_drawing(coord_action_map):
    result: DefaultDict[Coords, List[ops.Command]] = defaultdict(list)

    impositions: DefaultDict[Coords, List[ops.Command]] = defaultdict(list)

    r1, c1, r2, c2 = (None,) * 4
    for coords, actions in coord_action_map.items():
        for action in actions:
            if isinstance(action, ops.DrawBoxBorderOp):
                r1, c1 = action.top_left_point
                r2, c2 = action.bottom_right_point

                for r in range(r1, r2 + 1):
                    impositions[(r, c1)].append(
                        ops.ImposeFormat.with_format(action.left_format)
                    )
                    impositions[(r, c2)].append(
                        ops.ImposeFormat.with_format(action.right_format)
                    )
                for c in range(c1, c2 + 1):
                    impositions[(r1, c)].append(
                        ops.ImposeFormat.with_format(action.top_format)
                    )
                    impositions[(r2, c)].append(
                        ops.ImposeFormat.with_format(action.bottom_format)
                    )
            else:
                result[coords].append(action)

    def impose_to_target(row, col, format_):
        target_cell = (row, col)
        if not any(hasattr(op, 'data') for op in result[target_cell]):
            result[target_cell].append(ops.Write.with_data(None))
        result[target_cell].append(ops.ImposeFormat.with_format(format_))

    for coords, imposition in impositions.items():
        if not any(hasattr(op, 'data') for op in result[coords]):
            result[coords].append(ops.Write.with_data(None))

        merge_ops = [
            op
            for op in result[coords]
            if isinstance(op, ops.MergeWriteOp)
        ]
        if merge_ops:
            # There is a bug that makes right border of merged cells to not be written
            # So I have to highlight next cell's left border to make it look like a complete box
            max_width = max(
                op.size
                for op in merge_ops
            )

            impose_to_target(coords[0], coords[1] + max_width + 1, F.left_border)

            # Apparently Excel requires all cells to have the border whereas LibreOffice and Google Docs do fine
            # with just one merged cell with borders... until you unmerge them, that is.
            # Glorious copypasta ensues
            for column in range(coords[1], coords[1] + max_width + 1):
                if coords[0] < 1:
                    warn("A row above is required to impose top border to a merged cell group.")
                    break
                impose_to_target(coords[0] - 1, column, F.bottom_border)

                if r1 == r2:
                    impose_to_target(coords[0] + 1, column, F.top_border)

        result[coords].extend(imposition)

    return result


def _override_applier(coord_action_map):
    result = []

    def trigger_cell_dsl_error(message, exc=None):
        raise CellDSLError(message, action=action) from exc

    for key, actions in coord_action_map.items():
        transformed_actions = []

        sorted_actions = sorted(
            actions,
            key=lambda x: not isinstance(x, (ops.ImposeFormatOp, ops.OverrideFormatOp))
        )
        imposition_focus = {}
        override_focus = None
        for action in sorted_actions:
            if isinstance(action, ops.ImposeFormatOp):
                imposition_focus.update(action.format_)
            elif isinstance(action, ops.OverrideFormatOp):
                if override_focus is not None:
                    trigger_cell_dsl_error(f"There's already an OverrideFormat for cell {key}")
                override_focus = action
            else:
                if isinstance(action, (ops.WriteOp, ops.MergeWriteOp)):
                    if override_focus:
                        action = evolve(action, set_format=override_focus.format_)
                    elif imposition_focus:
                        action = action.with_format(imposition_focus)
                elif isinstance(action, ops.WriteRichOp):
                    if override_focus or imposition_focus:
                        # TODO: Figure out how to deal with impositions on those (should they be applied
                        #   to every text run?)
                        raise NotImplementedError('WriteRich is not supported with format impositions yet')
                transformed_actions.append(action)
        result.extend(
            (key, action)
            for action in transformed_actions
        )

    return result


def _process_chain(action_chain, initial_row, initial_col):
    coord_action_map, save_points = _process_movement(action_chain, initial_row, initial_col)
    coord_action_map, ref_array = _inject_coords(coord_action_map, save_points)
    coord_action_map = _introduce_ref_arrays(coord_action_map, ref_array)
    expanded_action_map = _expand_drawing(coord_action_map)
    coord_action_pairs = _override_applier(expanded_action_map)

    # Sort actions to go left-to-right, top-to-bottom
    coord_action_pairs.sort(key=lambda x: x[0][1])
    coord_action_pairs.sort(key=lambda x: x[0][0])

    return coord_action_pairs, save_points


@attrs(auto_attribs=True)
class ExecutorHelper(object):
    """A special object that performs some preprocessing of the commands when `commit` is called
    and stores the actions to be executed in `action_chain`"""
    action_chain: List[ops.Command] = Factory(list)

    def commit(self, chain: CommitTypes):
        """
        Commit this chain to `action_chain`.

        Args:
            chain: A tree with commands, int, str, dict or None..

        Notes:
            `int` acts like a shortcut for one or several `MoveOp` commands. Look at how
            the numbers correspond to the step direction, same as using a numpad.

            .. code:: text

                7 8 9 ↖ ↑ ↗
                4 5 6 ← . →
                1 2 3 ↙ ↓ ↘

        Examples:
            >>> E1 = ExecutorHelper()
            >>> E1.commit([
            ...     Write.with_data('Alpha'), 943, # Write alpha, follow by moving ↗←↘ --> →
            ...     [Write.with_data('Beta')], # Nested sequences are flattened
            ...     "Gamma", 2, # Lone strings are short forms of WriteOp.with_data(...)
            ...     # Dictionaries followed by a string is a short form of
            ...     # WriteOp.with_data(str).with_format(...)
            ...     F.default_font, "Delta", 2,
            ...     # Consecutive short forms of WriteOp are merged into WriteRich
            ...     F.default_font, F.center, "Epsilon ",
            ...     F.default_font, F.bold, " Eta",
            ...     F.italic, "!",
            ...     None # None is skipped
            ... ])
            >>> E2 = ExecutorHelper()
            >>> E2.commit([
            ...     Write.with_data('Alpha'),
            ...     Move.c(1).r(0), # →
            ...     Write.with_data('Beta'),
            ...     Write.with_data("Gamma"),
            ...     Move.r(1),
            ...     Write.with_data("Delta").with_format(F.default_font),
            ...     Move.r(1),
            ...     WriteRich.with_data("Epsilon ").with_format(F.default_font | F.center).then(
            ...         WriteRich.with_data(" Eta").with_format(F.default_font | F.bold)
            ...     ).then(
            ...         # .`with_format` implicitly merges with `F.default_font`, more specifically
            ...         #  `Format.FALLBACK_FONT`
            ...         WriteRich.with_data("!").with_format(F.default_font | F.italic)
            ...     )
            ... ])
            >>> E1.action_chain == E2.action_chain
            True
        """
        if not isinstance(chain, Iterable):
            chain = [chain]

        prev_format = None
        write_op_chain = []

        for subchain in chain + [None]:
            subchain_type = type(subchain)

            if prev_format is not None and subchain_type not in (FormatAsDictForm, FormatDict, WriteDataShortForm):
                raise CellDSLError(f"Format shortcut must be followed by a format or string shortcut: {prev_format}"
                                   f" {subchain}")
            elif subchain_type is WriteDataShortForm:
                action = ops.Write.with_data(subchain)
                if prev_format:
                    action = action.with_format(prev_format)
                write_op_chain.append(action)
                prev_format = None
                continue
            elif subchain_type in (FormatAsDictForm, FormatDict):
                if prev_format:
                    prev_format |= subchain
                else:
                    if subchain_type is FormatAsDictForm:
                        prev_format = FormatDict(cast(FormatAsDictForm, subchain))
                    else:
                        prev_format = subchain
                continue

            if len(write_op_chain) > 1:
                # WriteRich form
                to_add = WriteRich.with_data(write_op_chain[0].data).with_format(write_op_chain[0].format_)
                for segment in write_op_chain[1:]:
                    to_add = to_add.then(WriteRich.with_data(segment.data).with_format(segment.format_))
                self.action_chain.append(to_add)
                write_op_chain.clear()
            elif len(write_op_chain) == 1:
                # WriteOp form
                self.action_chain.extend(write_op_chain)
                write_op_chain.clear()

            if subchain_type is MovementShortForm:
                delta_row, delta_col = 0, 0
                for direction in str(subchain):
                    delta_row += direction in '123' and 1 or direction in '789' and -1 or 0
                    delta_col += direction in '369' and 1 or direction in '147' and -1 or 0
                self.action_chain.append(ops.Move.r(delta_row).c(delta_col))
            elif isinstance(subchain, ops.Command):
                self.action_chain.append(subchain)
            elif isinstance(subchain, Iterable):
                self.commit(subchain)
            elif subchain is None:
                pass
            else:
                raise CellDSLError(f'Cannot process this type: {subchain_type}')


@contextmanager
def cell_dsl_context(
        target: WorksheetTriplet,
        initial_row: int = 0,
        initial_col: int = 0,
        stat_receiver: Optional[StatReceiver] = None
) -> ContextManager[ExecutorHelper]:
    """Creates a context inside which you can perform writes and change cells in `target` in an arbitrary order,
    starting at `initial_row` and `initial_col`.

    After it exits, it will execute those changes and send stats to `stat_receiver`.

    Parameters:
        target: Target WriterPair to apply changes with
        stat_receiver:
            This is a reference to a StatReceiver object that will receive all data about visited cells and save points.
        initial_row: Starting row, zero-based
        initial_col: Starting column, zero-based

    Returns:
        ExecutorHelper:
            A special dummy class that commits changes via `commit` method.
    """
    helper = ExecutorHelper()
    yield helper

    coord_action_pairs, save_points = _process_chain(helper.action_chain, initial_row, initial_col)
    name_stack = deque()

    try:
        cell_dsl_context.hbreaks
    except AttributeError:
        cell_dsl_context.hbreaks = set()

    try:
        cell_dsl_context.vbreaks
    except AttributeError:
        cell_dsl_context.vbreaks = set()

    def trigger_execution_error(message):
        raise ExecutionCellDSLError(message, action_num, None, name_stack, action, save_points)

    try:
        for action_num, (coords, action) in enumerate(coord_action_pairs):
            action_type = type(action)
            if action_type is ops.SubmitHPagebreakOp:
                cell_dsl_context.hbreaks.add(coords[0])
            elif action_type is ops.SubmitVPagebreakOp:
                cell_dsl_context.vbreaks.add(coords[1])
            elif action_type is ops.ApplyPagebreaksOp:
                target.ws.set_h_pagebreaks([*cell_dsl_context.hbreaks])
                target.ws.set_v_pagebreaks([*cell_dsl_context.vbreaks])
                cell_dsl_context.hbreaks.clear()
                cell_dsl_context.vbreaks.clear()
            elif action_type is ops.SectionBeginOp:
                name_stack.append(action.name)
            elif action_type is ops.SectionEndOp:
                name_stack.pop()
            elif isinstance(action, ExecutableCommand):
                action.execute(target, coords)
            else:
                raise TypeError(f'Unknown action of type {type(action)}: {action}')
    except CellDSLError as e:
        trigger_execution_error(e.message)

    if stat_receiver is not None:
        # We need to insert an implicit first action of moving over to
        #   starting position
        coord_action_pairs.insert(
            0, ((initial_row, initial_col), None)
        )
        stat_receiver.initial_row = initial_row
        stat_receiver.initial_col = initial_col
        stat_receiver.coord_pairs = coord_action_pairs
        stat_receiver.save_points = save_points


@contextmanager
def dummy_cell_dsl_context(
        initial_row: int = 0,
        initial_col: int = 0,
        stat_receiver: StatReceiver = None
):
    """A version of `cell_dsl_context` that does not actually execute its actions, used in testing and debug instead."""
    helper = ExecutorHelper()
    yield helper

    coord_action_pairs, save_points = _process_chain(helper.action_chain, initial_row, initial_col)
    stat_receiver.initial_row = initial_row
    stat_receiver.initial_col = initial_col
    stat_receiver.coord_pairs = coord_action_pairs
    stat_receiver.save_points = save_points
