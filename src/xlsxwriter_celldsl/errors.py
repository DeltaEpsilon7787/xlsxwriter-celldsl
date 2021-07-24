from itertools import groupby
from pprint import pformat


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


class ExecutionCellDSLError(CellDSLError):
    """A Cell DSL error triggered by an exception during execution."""
