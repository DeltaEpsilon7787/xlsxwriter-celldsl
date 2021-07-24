from collections import defaultdict, deque
from itertools import chain as itchain
from operator import itemgetter
from typing import DefaultDict, Dict, Iterable, Iterator, List, Optional, Tuple, Union, cast
from warnings import warn

from attr import Factory, attrib, attrs, evolve
from xlsxwriter.utility import xl_range_abs

from . import ops
from .errors import CellDSLError, ExecutionCellDSLError, MovementCellDSLError
from .formats import FormatDict, FormatsNamespace as F
from .utils import WorksheetTriplet

MovementShortForm = int
WriteDataShortForm = str
FormatAsDictForm = dict

CommitTypes = Union[
    Iterable['CommitTypes'],
    ops.classes.Command,
    MovementShortForm,
    WriteDataShortForm,
    FormatAsDictForm,
    FormatDict,
    None
]

CoordActionPair = Tuple[ops.traits.Coords, ops.classes.Command]


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
    coord_pairs: List[Tuple[ops.traits.Coords, ops.classes.Command]] = attrib(init=False)
    save_points: Dict[str, ops.traits.Coords] = attrib(init=False, factory=dict)

    @property
    def is_null(self):
        return len(self.coord_pairs) < 2

    @property
    def _coord_iter(self) -> Iterator[ops.traits.Coords]:
        return cast(Iterator[ops.traits.Coords], map(itemgetter(0), self.coord_pairs))

    @property
    def max_row(self) -> int:
        return cast(int, max(self._coord_iter, key=itemgetter(0))[0])

    @property
    def max_col(self) -> int:
        return cast(int, max(self._coord_iter, key=itemgetter(1))[1])

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
    def max_coords(self) -> ops.traits.Coords:
        return self.max_row, self.max_col

    @property
    def initial_coords(self) -> ops.traits.Coords:
        return self.initial_row, self.initial_col


@attrs(auto_attribs=True)
class _SetNameStack(object):
    stack: deque


@attrs(auto_attribs=True)
class ExecutorHelper(object):
    """A special object that performs some preprocessing of the commands when `commit` is called
    and stores the actions to be executed in `action_chain`"""
    action_chain: List[ops.classes.Command] = Factory(list)

    def commit(self, chain: CommitTypes):
        """
        Commit this `chain` to `action_chain`.

        Notes:
            `int` acts like a shortcut for one or several `MoveOp` commands. Look at how
            the numbers correspond to the step direction, same as using a numpad.

            .. code:: text

                7 8 9 ↖ ↑ ↗
                4 5 6 ← . →
                1 2 3 ↙ ↓ ↘

        Examples:
            >>> from xlsxwriter_celldsl.ops import Write, Move, WriteRich
            >>> from xlsxwriter_celldsl.formats import FormatsNamespace as F
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
            ...     # Ending the short form with a format will set this cell to this format
            ...     F.wrapped
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
            ...         # `with_format` implicitly merges with `F.default_font`, more specifically
            ...         #  `Format.FALLBACK_FONT`
            ...         WriteRich.with_data("!").with_format(F.default_font | F.italic)
            ...         # Cell formats here are an exception to implicit merges
            ...         #  since it's not used to set the text appearance, but the cell's
            ...         .with_cell_format(F.wrapped)
            ...     )
            ... ])
            >>> E1.action_chain == E2.action_chain
            True
        """
        if not isinstance(chain, Iterable):
            chain = [chain]

        prev_format = None
        write_op_chain = []

        for subchain in itchain(chain, [None]):
            subchain_type = type(subchain)

            if subchain_type is WriteDataShortForm:
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
                to_add = ops.WriteRich.with_data(write_op_chain[0].data).with_format(write_op_chain[0].format_)

                for segment in write_op_chain[1:]:
                    to_add = to_add.then(ops.WriteRich.with_data(segment.data).with_format(segment.format_))

                if prev_format:
                    # Ended with a format, Cell Format variation
                    to_add = to_add.with_cell_format(prev_format)
                self.action_chain.append(to_add)
                write_op_chain.clear()
            elif len(write_op_chain) == 1:
                # WriteOp form
                self.action_chain.extend(write_op_chain)
                write_op_chain.clear()

            if subchain is None:
                pass
            elif subchain_type is MovementShortForm:
                delta_row, delta_col = 0, 0
                for direction in str(subchain):
                    delta_row += direction in '123' and 1 or direction in '789' and -1 or 0
                    delta_col += direction in '369' and 1 or direction in '147' and -1 or 0
                self.action_chain.append(ops.Move.r(delta_row).c(delta_col))
            elif isinstance(subchain, ops.classes.Command):
                self.action_chain.append(subchain)
            elif isinstance(subchain, Iterable):
                self.commit(subchain)
            else:
                raise CellDSLError(f'Cannot process this type: {subchain_type}, {subchain}')


class CellDSLContext(object):
    """
    A context manager inside which you can perform writes and change cells in `target` in an arbitrary order,
    starting at `initial_row` and `initial_col`, even writing multiple times to a cell if `overwrites_ok`.

    After it exits, it will execute those changes and send stats to `stat_receiver`.

    This context manager returns an :class:`ExecutorHelper` object, use :func:`commit` method
    to submit operations to be executed.

    Parameters:
        target: Target WorksheetTriplet to apply changes with
        stat_receiver:
            This is a reference to a StatReceiver object that will receive all data about visited cells and save points.
        initial_row: Starting row, zero-based
        initial_col: Starting column, zero-based
        overwrites_ok:
            If True, it is expected that there may be several writes with different data to the same cell, which is
            ordinarily a sign of a bug since the result would be ambiguous.
            If False, overwrites raise :class:`ExecutionCellDSLError`.
    """

    def __init__(
            self,
            target: WorksheetTriplet,
            initial_row: int = 0,
            initial_col: int = 0,
            stat_receiver: Optional[StatReceiver] = None,
            overwrites_ok: bool = False,
    ):
        self.target = target
        self.initial_row = initial_row
        self.initial_col = initial_col
        self.stat_receiver = stat_receiver
        self.overwrites_ok = overwrites_ok

        self.helper: ExecutorHelper = ExecutorHelper()

    @staticmethod
    def _process_movement(action_list, row, col) -> \
            Tuple[DefaultDict[ops.traits.Coords, List[ops.classes.Command]], Dict[str, ops.traits.Coords]]:
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
            if action_type is ops.classes.LoadOp:
                try:
                    row, col = save_points[action.point_name]
                except KeyError as e:
                    trigger_movement_error(f'Save point {action.point_name} does not exist.', e)
                visited.append((row, col))
            elif action_type is ops.classes.StackLoadOp:
                try:
                    row, col = save_stack.pop()
                except IndexError as e:
                    trigger_movement_error(f'Save stack is empty.', e)
                visited.append((row, col))
            elif action_type is ops.classes.SaveOp:
                save_points[action.point_name] = (row, col)
            elif action_type is ops.classes.StackSaveOp:
                save_stack.append((row, col))
            elif action_type is ops.classes.MoveOp:
                row += action.row
                col += action.col
                visited.append((row, col))
            elif action_type is ops.classes.AtCellOp:
                row = action.row
                col = action.col
                visited.append((row, col))
            elif action_type is ops.classes.BacktrackCellOp:
                try:
                    for _ in range(action.n + 1):
                        row, col = visited.pop()
                except IndexError as e:
                    trigger_movement_error(f'Could not backtrack {action.n} cells.', e)
            elif action_type is ops.classes.SectionBeginOp:
                name_stack.append(action.name)
            elif action_type is ops.classes.SectionEndOp:
                name_stack.pop()
            else:
                if isinstance(action, ops.traits.Range):
                    if isinstance(action.top_left_point, int):
                        if action.top_left_point > 0:
                            try:
                                action = action.with_top_left(visited[-action.top_left_point - 1])
                            except IndexError as e:
                                trigger_movement_error(
                                    f'Top left corner would use {action.top_left_point} last visited cell, but only '
                                    f'{len(visited)} cells have been visited',
                                    e
                                )
                        elif action.top_left_point < 0:
                            try:
                                action = action.with_top_left(save_stack[action.top_left_point])
                            except IndexError as e:
                                trigger_movement_error(
                                    f'Top left corner would look {-action.top_left_point} positions'
                                    f'up the save stack, but there is only '
                                    f'{len(save_stack)} saves',
                                    e
                                )
                        else:
                            action = action.with_top_left((row, col))
                    if isinstance(action.bottom_right_point, int):
                        if action.bottom_right_point > 0:
                            try:
                                action = action.with_bottom_right(visited[-action.bottom_right_point - 1])
                            except IndexError as e:
                                trigger_movement_error(
                                    f'Bottom right corner would use {action.bottom_right_point} '
                                    f'last visited cell, but only {len(visited)} cells have been visited',
                                    e
                                )
                        elif action.bottom_right_point < 0:
                            try:
                                action = action.with_bottom_right(save_stack[action.bottom_right_point])
                            except IndexError as e:
                                trigger_movement_error(
                                    f'Bottom right corner would look {-action.bottom_right_point} positions'
                                    f'up the save stack, but there is only '
                                    f'{len(save_stack)} saves',
                                    e
                                )
                        else:
                            action = action.with_bottom_right((row, col))

                result[(row, col)].append(action.absorb_name_stack_data(name_stack))

            # 2^20 and 2^14 are Excel limits for the amount of row and columns respectively.
            if row not in range(0, 2 ** 20) or col not in range(0, 2 ** 14):
                trigger_movement_error(f'Illegal coords have been reached, this is not allowed.')

        if len(name_stack) > 0:
            raise trigger_movement_error(f'Name stack is not empty, every SectionBegin must be matched with SectionEnd')

        return result, save_points

    @staticmethod
    def _inject_coords(coord_action_map, save_points) -> \
            Tuple[DefaultDict[ops.traits.Coords, List[ops.classes.Command]], Dict[str, str]]:
        result = defaultdict(list)
        ref_array = {}

        def trigger_cell_dsl_error(message, exc=None):
            raise CellDSLError(message, action=action, save_points=save_points) from exc

        for coords, actions in coord_action_map.items():
            for action in actions:
                if isinstance(action, ops.traits.Range):
                    if isinstance(action.top_left_point, str):
                        try:
                            action = action.with_top_left(save_points[action.top_left_point])
                        except KeyError as e:
                            trigger_cell_dsl_error(
                                f"Tried to use a save point named {action.top_left_point} "
                                f"for top left corner, but it doesn't exist",
                                e
                            )
                    if isinstance(action.bottom_right_point, str):
                        try:
                            action = action.with_bottom_right(save_points[action.bottom_right_point])
                        except KeyError as e:
                            trigger_cell_dsl_error(
                                f"Tried to use a save point named {action.bottom_right_point} "
                                f"for bottom right corner, but it doesn't exist",
                                e
                            )

                if isinstance(action, ops.classes.RefArrayOp):
                    c = xl_range_abs(*action.top_left_point, *action.bottom_right_point)
                    ref_array[action.point_name] = f'{c}'
                    continue

                result[coords].append(action)

        return result, ref_array

    @staticmethod
    def _introduce_ref_arrays(coord_action_map, ref_array):
        result: DefaultDict[ops.traits.Coords, List[ops.classes.Command]] = defaultdict(list)

        for coords, actions in coord_action_map.items():
            for action in actions:
                if isinstance(action, ops.traits.ForwardRef):
                    action = action.inject_refs(ref_array)

                result[coords].append(action)

        return result

    @staticmethod
    def _expand_drawing(coord_action_map):
        result: DefaultDict[ops.traits.Coords, List[ops.classes.Command]] = defaultdict(list)

        impositions: DefaultDict[ops.traits.Coords, List[ops.classes.Command]] = defaultdict(list)

        r1, c1, r2, c2 = (None,) * 4
        for coords, actions in coord_action_map.items():
            for action in actions:
                if isinstance(action, ops.classes.DrawBoxBorderOp):
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
                if isinstance(op, ops.classes.MergeWriteOp)
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

    @staticmethod
    def _override_applier(coord_action_map):
        result = []

        def trigger_cell_dsl_error(message, exc=None):
            raise CellDSLError(message, action=action) from exc

        for key, actions in coord_action_map.items():
            transformed_actions = []

            sorted_actions = sorted(
                actions,
                key=lambda x: not isinstance(x, (ops.classes.ImposeFormatOp, ops.classes.OverrideFormatOp))
            )
            imposition_focus = FormatDict()
            override_focus = None
            for action in sorted_actions:
                if isinstance(action, ops.classes.ImposeFormatOp):
                    imposition_focus |= action.format_
                elif isinstance(action, ops.classes.OverrideFormatOp):
                    if override_focus is not None:
                        trigger_cell_dsl_error(f"There's already an OverrideFormat for cell {key}")
                    override_focus = action
                else:
                    if isinstance(action, (ops.classes.WriteOp, ops.classes.MergeWriteOp)):
                        if override_focus:
                            action = evolve(action, set_format=override_focus.format_)
                        elif imposition_focus:
                            action = action.with_format(imposition_focus)
                    elif isinstance(action, ops.classes.WriteRichOp):
                        if override_focus:
                            action = action.with_cell_format(override_focus.format_)
                        elif imposition_focus:
                            if action.cell_format:
                                action = action.with_cell_format(FormatDict(action.cell_format | imposition_focus))
                            else:
                                action = action.with_cell_format(FormatDict(imposition_focus.copy()))
                    transformed_actions.append(action)
            result.extend(
                (key, action)
                for action in transformed_actions
            )

        return result

    @staticmethod
    def _process_chain(action_chain, initial_row, initial_col) -> \
            Tuple[List[CoordActionPair], Dict[str, ops.traits.Coords]]:
        coord_action_map, save_points = CellDSLContext._process_movement(action_chain, initial_row, initial_col)
        coord_action_map, ref_array = CellDSLContext._inject_coords(coord_action_map, save_points)
        coord_action_map = CellDSLContext._introduce_ref_arrays(coord_action_map, ref_array)
        expanded_action_map = CellDSLContext._expand_drawing(coord_action_map)
        coord_action_pairs = CellDSLContext._override_applier(expanded_action_map)

        # Sort actions to go left-to-right, top-to-bottom
        coord_action_pairs.sort(key=lambda x: x[0][1])
        coord_action_pairs.sort(key=lambda x: x[0][0])

        return coord_action_pairs, save_points

    def __enter__(self):
        return self.helper

    def __exit__(self, exc_type, exc_val, exc_tb):
        coord_action_pairs, save_points = self._process_chain(
            self.helper.action_chain,
            self.initial_row, self.initial_col
        )

        try:
            cell_dsl_context.hbreaks
        except AttributeError:
            cell_dsl_context.hbreaks = set()

        try:
            cell_dsl_context.vbreaks
        except AttributeError:
            cell_dsl_context.vbreaks = set()

        def trigger_execution_error(message):
            raise ExecutionCellDSLError(message, action_num, None, action.NAME_STACK_DATA, action, save_points)

        override_tracking = {}

        try:
            for action_num, (coords, action) in enumerate(coord_action_pairs):
                action_type = type(action)

                if action_type is ops.classes.SubmitHPagebreakOp:
                    cell_dsl_context.hbreaks.add(coords[0])
                elif action_type is ops.classes.SubmitVPagebreakOp:
                    cell_dsl_context.vbreaks.add(coords[1])
                elif action_type is ops.classes.ApplyPagebreaksOp:
                    self.target.ws.set_h_pagebreaks([*cell_dsl_context.hbreaks])
                    self.target.ws.set_v_pagebreaks([*cell_dsl_context.vbreaks])
                    cell_dsl_context.hbreaks.clear()
                    cell_dsl_context.vbreaks.clear()
                else:
                    if not self.overwrites_ok and action.OVERWRITE_SENSITIVE:
                        if coords in override_tracking:
                            if action != override_tracking[coords]:
                                raise ExecutionCellDSLError(f'Overwrite has occurred at {coords}.')
                            continue
                        override_tracking[coords] = action
                    if isinstance(action, ops.traits.ExecutableCommand):
                        action.execute(self.target, coords)
                    else:
                        raise TypeError(f'Unknown action of type {type(action)}: {action}')
        except CellDSLError as e:
            trigger_execution_error(e.message)
        except Exception as e:
            raise ExecutionCellDSLError(
                'Uncaught exception',
                action_num,
                None,
                action.NAME_STACK_DATA,
                action,
                save_points
            ) from e

        if self.stat_receiver is not None:
            # We need to insert an implicit first action of moving over to
            #   starting position
            coord_action_pairs.insert(
                0, ((self.initial_row, self.initial_col), ops.Move)
            )
            self.stat_receiver.initial_row = self.initial_row
            self.stat_receiver.initial_col = self.initial_col
            self.stat_receiver.coord_pairs = coord_action_pairs
            self.stat_receiver.save_points = save_points


cell_dsl_context = CellDSLContext
