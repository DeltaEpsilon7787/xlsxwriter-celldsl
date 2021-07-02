"""Cell DSL implementation for XlsxWriter, providing ways to index cells using relative movement and giving names to
various key coordinates, overall allowing one to write Excel table generating code by imagining how they would do it
manually and translating the actions into commands in `ops` module, after `commit`ing them inside `cell_dsl_context`,
after which stats can be transmitted to a `StatReceiver` and used for further writing, all the while constructing
text formats by conjunction `FormatDict`s and using some default formats from `FormatsNamespace`."""

from . import cell_dsl, formats, ops, utils

from .cell_dsl import StatReceiver, cell_dsl_context
from .utils import WorkbookPair

__all__ = ['cell_dsl', 'formats', 'ops', 'utils', 'StatReceiver', 'cell_dsl_context', 'WorkbookPair']
