Cell DSL API
===========================

Traits
------
Traits represent abstract properties that various operations may have

.. currentmodule:: xlsxwriter_celldsl.traits
.. autosummary::
   :nosignatures:

   Data
   DataType
   Format
   RelativePosition
   AbsolutePosition
   FractionalSize
   CardinalSize
   Range
   NamedPoint
   ForwardRef
   Options
   ExecutableCommand

.. automodule:: xlsxwriter_celldsl.traits
   :members:

.. _api-ops:

Operations
----------
Operations are the commands that are executed after they're committed

.. currentmodule:: xlsxwriter_celldsl.ops
.. autosummary::
   :nosignatures:

   MoveOp
   AtCellOp
   BacktrackCellOp
   StackSaveOp
   StackLoadOp
   LoadOp
   SaveOp
   RefArrayOp
   SectionBeginOp
   SectionEndOp
   WriteOp
   MergeWriteOp
   WriteRichOp
   ImposeFormatOp
   OverrideFormatOp
   DrawBoxBorderOp
   DefineNamedRangeOp
   SetRowHeightOp
   SetColumnWidthOp
   SubmitHPagebreakOp
   SubmitVPagebreakOp
   ApplyPagebreaksOp
   AddCommentOp
   AddChartOp
   AddConditionalFormatOp
   AddImageOp

.. automodule:: xlsxwriter_celldsl.ops
   :members:
   :show-inheritance:

Utils
-----
Various helpful objects and functions

.. automodule:: xlsxwriter_celldsl.utils
   :members:

Formats
-------
Various utilities to deal with formats

.. automodule:: xlsxwriter_celldsl.formats
   :members:

Cell DSL module
---------------

.. automodule:: xlsxwriter_celldsl.cell_dsl
   :members:
