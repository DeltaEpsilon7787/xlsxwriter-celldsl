Intermediate usage
==================

.. py:currentmodule:: xlsxwriter_celldsl.ops.classes

Advanced movement and Stack saves
---------------------------------
Cell DSL also comes with more unusual forms of save points and movement.

.. autoclass:: AtCellOp
   :noindex:

.. autoclass:: BacktrackCellOp
   :noindex:

.. autoclass:: StackSaveOp
   :noindex:

.. autoclass:: StackLoadOp
   :noindex:

AtCellOp is used in those rare cases where there is a necessity to jump to a very specific location,
and it's not the initial offset for whatever the current context is writing.

.. warning::
    Do not use AtCellOp to init the imaginary cursor, use `initial_row` and `initial_col` arguments on
    :func:`cell_dsl_context <xlsxwriter_celldsl.cell_dsl.cell_dsl_context>` instead.

BacktrackCellOp is an unusual form of movement which, instead of going a specific amount of cells away
or going to some known cell, jumps back to a cell that was visited some time ago. This is mostly useful
in cases where you have a structure generating function that performs some unwanted moves at the end.

For example, in case of Cell DSL own utility function :func:`row_chain <xlsxwriter_celldsl.utils.row_chain>`,
the unwanted last movement may be the jump back to starting position. Perhaps you want to continue
writing to the row outside of the `row_chain`. To do so, you want to return back to the last position
it wrote to, thus `BacktrackCellOp` proves useful here.

StackSave and StackLoad utilize a separate save point system which, instead of using a dictionary
of string names to coordinates, uses a stack of coordinates, without any names. This is useful
in generating functions in order to avoid polluting the save points of the host
function. For example, aforementioned `row_chain` uses those operations in order to
return back to starting position without ever using :class:`SaveOp`.

Format imposition and override
------------------------------

In some cases it may be useful to apply changes to `formats` of writing operations after they're
submitted. An example of such a situation is when wants to highlight an arbitrary cell, which,
in case of raw XlsxWriter, would require to calculating this cell in advance and then awkwardly
merging the format at its write call with the format of the highlight.

Cell DSL provides a way to perform this operation.

.. autoclass:: ImposeFormatOp
   :noindex:

.. autoclass:: OverrideFormatOp
   :noindex:

Advanced writes
---------------

Besides basic :class:`WriteOp`, Cell DSL also comes with two additional ways to write to a cell.

.. autoclass:: WriteRichOp
   :noindex:

.. autoclass:: MergeWriteOp
   :noindex:

WriteRichOp is specifically useful to write strings that have several formats into a single cell and
also comes with a short form in :func:`commit <xlsxwriter_celldsl.cell_dsl.ExecutorHelper.commit>` method.

MergeWriteOp is a strange variation of a    writing operation that has a :class:`Range <xlsxwriter_celldsl.ops.traits.Range>`
trait, mirroring the behavior of :func:`merge_range` method in XlsxWriter.

Ranged commands
---------------
There are several operations in Cell DSL which target a range of cells, all of which have quirks in what they do.

.. autoclass:: xlsxwriter_celldsl.ops.traits.Range
   :noindex:

The range trait itself has a fairly complicated set of rules regarding its value types.

.. autoclass:: DefineNamedRangeOp
   :noindex:

Just as using named ranges is good practice in Excel, you may also leverage their power in Cell DSL
in situations where you need to refer to a literal cell range by name, such as in formulas.

However, an exception to this exists where a literal cell range must be by value only: charts.
Cell DSL provides a separate operation for charts and related commands specifically.

.. autoclass:: RefArrayOp
   :noindex:

Operations with a trait of :class:`ForwardRef <xlsxwriter_celldsl.ops.traits.ForwardRef>` receive
a literal cell range into `resolved_refs` at the time of their execution.
This is only relevant for custom commands that want to take advantage of this functionality.

Charts
------
Unlike many other operations, charts are significantly more involved to work with in both XlsxWriter and Cell DSL.

.. autoclass:: AddChartOp
   :noindex:

As Cell DSL doesn't execute operations until exiting the context,
chart methods may not be called during submission. In order to mitigate this,
chart operations come with their own action chains,
but there is no need to create a separate operation set for them as long as only methods are called.

Every instance of AddChartOp is associated with a specific type of Chart and a specific XlsxWriter class.
Those classes provide various methods depending on the type of the chart.

In order to queue up methods to be called during Cell DSL execution stage, use `target` attribute of AddChartOp.
This attribute will contains an object that mimics the associated chart class and
allows you to call various chart methods as if it was a real instance of that chart class.

Here is an example of it in action::

    from xlsxwriter_celldsl.ops import AddBarChart, RefArray

    with cell_dsl_context(ws_triplet) as E:
        E.commit([
            RefArray.top_left((0, 0)).bottom_right((0, 3)).at('some ref array'),
            AddBarChart.do([
                AddBarChart.target.add_series({'values': RefArray.at('some ref array')}),
                AddBarChart.target.add_series({'values': '=TestSheet!$A$2:$C$7'})
            ]),
        ])

This will ensure that :func:`add_series` of XlsxWriter is called during execution.
Values of type RefArrayOp will automatically be replaced with a literal cell range string,
in this case this string will be ``'=SheetName!$A$1:$D$1'``,
because coordinates of ``A1`` are ``(0, 0)`` and of ``D1`` are ``(0, 3)``.

.. note::
    Substitution functionality is unique to charts, custom commands will have to reinvent the wheel.