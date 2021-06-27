.. py:currentmodule:: xlsxwriter_celldsl.utils

Basic usage
===========

Workbook - FormatHandler pair
-----------------------------

To start using Cell DSL, you first need to generate a :class:`WorkbookPair`, which is a simple object that
couples together a :ref:`Workbook <workbook>` object with
:class:`FormatHandler <xlsxwriter_celldsl.formats.FormatHandler>`.
Only one such object should exist for a single Workbook,
this is needed to ensure formats added to the worksheet are kept to a minimum.

To create a Workbook pair, along with its associated FormatHandler, simply use :func:`from_wb` class method::

    from xlsxwriter_celldsl import WorkbookPair

    # Assume wb is a xlsxwriter.Workbook object
    wb_pair = WorkbookPair.from_wb(wb)

Worksheet - FormatHandler - Workbook triplet
--------------------------------------------
Following WorkbookPair, you want to then create a :class:`WorksheetTriplet`,
which uses the data from a host WorkbookPair and creates a :ref:`Worksheet <worksheet>` associated with
the host Workbook and FormatHandler. This is the main data object used by this library.

To create a WorksheetTriplet, call :func:`WorkbookPair.add_worksheet` which mirrors :func:`add_worksheet` method::

    ws_triplet = wb_pair.add_worksheet('SheetName')

Cell DSL entry point
--------------------
.. py:currentmodule:: xlsxwriter_celldsl.cell_dsl

After creating a WorksheetTriplet, you are ready to start using Cell DSL.
The entry point to the library is a :py:ref:`context manager <context-managers>` :func:`cell_dsl_context`.
Returning an :class:`ExecutorHelper` object, it creates a context within which you can :func:`ExecutorHelper.commit`
commands to be executed within the created Worksheet.

It has the following signature:

.. autofunction:: cell_dsl_context
   :noindex:

Before entering the context, you should decide whether you want to store a few stats or not.
To do that, you need to instantiate a :class:`StatReceiver` object.

.. autoclass:: StatReceiver
   :noindex:

After exiting the context, this object will be populated with respective values.

The recommended way to do all of the above is as follows::

    # Do if necessary
    from xlsxwriter_celldsl.cell_dsl import StatReceiver, cell_dsl_context
    stat_receiver = StatReceiver()

    with cell_dsl_context(ws_triplet, stat_receiver) as E:
        ...

Operations
----------
Cell DSL central objects are operations, all of which are stored in :ref:`ops module <api-ops>`.

Operations are accepted by :class:`ExecutorHelper` via its :func:`commit <ExecutorHelper.commit>` method.
Besides keeping track of submitted operations, this method also performs some preprocessing,
providing several short forms of common operations, namely for writing and movement.

:ref:`ops module <api-ops>` exports two types of names: operation classes and base operation instances.

Each instance has a verb-like name and its respective class is the same name ending with *Op*.
So, `ops.Write` is an instance of `ops.WriteOp`.

As `commit` accepts instances, not classes, this means you don't need to instantiate your own base instances.

Each operation is defined using `traits`, which provide the operation some parameters and
respective methods to modify those parameters.

Unless you have a really good reason, you should only parametrize an operation using those provided methods.
As all operations are immutable, the methods create a new instance of this operation with the related
parameter changed.

Thus, a script is a list of operations, and each operation is parametrized using chained method calls, such as::

    [
        Write
            .with_data("Example"),
        Move
            .r(1)
            .c(2),
        ...
    ]

Configure your linter as needed. The docstring of each operation gives a hint as to what it does and which methods
should be used to parametrize it.

Basic movement and basic write
------------------------------

The two most basic operations, an operation to move an imaginary cursor within the spreadsheet
and an operation to perform a write into the cell the cursor is at.

.. py:currentmodule:: xlsxwriter_celldsl.ops

.. autoclass:: MoveOp
   :noindex:

.. autoclass:: WriteOp
   :noindex:

.. py:currentmodule:: xlsxwriter_celldsl.cell_dsl.ExecutorHelper

Both operations also happen to have a short form, refer to the example code of `func`:`commit` method
for elaboration.

Save / Load
-----------

.. py:currentmodule:: xlsxwriter_celldsl.ops

But movement with just :class:`MoveOp` would be quite limited and, because it utilizes relative movement,
you do not know which position you're at unless you read the entire context beforehand.

Which is why Cell DSL comes with a powerful system of *save points*, which is a way to give current location
a name and be able to go back to it from anywhere later.

.. autoclass:: SaveOp
   :noindex:

.. autoclass:: LoadOp
   :noindex:

It is therefore a good practice to utilize save points as much as needed in order to give notable positions
a meaningful name, reducing the need to know the entire history of movement.

Another notable use of save point system is for testing code using Cell DSL because
one of the values :class:`StatReceiver <xlsxwriter_celldsl.cell_dsl>` gets is a mapping of names to coordinates.

Sectioning
----------

Errors are inevitable in code and, due to the nature of Cell DSL mode of operation, it is harder to
debug them because execution occurs upon exiting the cell_dsl_context, after all actions have been submitted.

Because actions are submitted using normal Python's methods, there is no way to track back which
line of code the exception occurs at. This is the curse of any system where the execution doesn't occur immediately,
and Cell DSL is no exception.

However, various data from the context of execution is still available for usage which may help in tracking
down the location of the bug.

One of the more explicit ways to do so is by using sectioning.

.. autoclass:: SectionBeginOp
   :noindex:

.. autoclass:: SectionEndOp
   :noindex:

These two operations allow giving various sections in the script a name. Think of them like named code blocks.
Code blocks may also be nested, same as sections. During execution, whenever `SectionBeginOp` is reached,
the name is pushed into a name stack or popped if `SectionEndOp` is reached.

By sectioning your script code into blocks, you reduce the amount of work needed to track down the offending action
because, in the exception message, the name stack will always be available for you.

In principle, if every single operation was a section, you'd be able always know the exact place an error occurs at.
However, doing so in practice would mean a lot of visual noise and is unacceptable.

A text stream parser type DSL would be able to imbue every token with tracking information, but Cell DSL
does not receive a text stream, but what is essentially already a parse tree, thus sectioning was a compromise.

Besides debugging purposes, sectioning may also be used to document your code by giving various sections a name,
however for this purpose using pure Python functions is more appropriate.

Custom formats
--------------
.. py:currentmodule:: xlsxwriter_celldsl.formats

Finally, it is time to talk about formats. Formats are essentially the main way to modify the outlook of
both text and cells in Excel. However, one annoying thing about them is that in XlsxWriter formats can only be
changed and applied during writing and some other specific writing operations. Worse still, formats
are treated as objects that are added into the Excel sheet and require additional care
in order to be useful.

Further exacerbating the issue is the fact that formats cannot be merged together or modified.
When working in Excel, one can easily apply changes like making text bold or changing the font without
having to rewrite it. By contrast, XlsxWriter will require you to manually transfer information about
the original format in the target cell, so, if you just want to make text bold in a specific cell,
you have to also make sure to specify the font name, size and other features.

Even though you are essentially just trying to compose two formats together, in reality you are
merging them manually and it's your responsibility to make sure information is carried all the way through.

Cell DSL removes the need to keep track of those details and utilizes its own special type
in order to make formats easily composable.

.. autoclass:: FormatDict
   :noindex:

Using this type, you can forget about keeping track of redundant format details and instead treat format objects
as a composition of smaller format traits.

All operations that utilize a format as one of the parameters will only accept FormatDict and plain dictionaries.
During execution, handling converting format dictionaries into actual format objects will be done implicitly
and in a way that creates the least amount of formats necessary, thus completely removing any need to deal with
raw format objects directly.

.. warning::
  There is an explicit base format used in all operations.
  All operations for which a custom format is not specified will use it.
  All custom formats are implicitly merged with this base format barring :class:`xlsxwriter_celldsl.ops.DrawBoxBorderOp`.
  This format can be configured using :func:`set_base_format` method of the specific operation.
  Do keep in mind that this changes it globally and should only be done once as such.

Operations that use formats inherit :class:`xlsxwriter_celldsl.traits.Format` trait.

.. autoclass:: xlsxwriter_celldsl.traits.Format
   :members:
   :noindex:

Additionally, Cell DSL provides a set of default formats, check :class:`FormatsNamespace` source code to see
what the formats are like.

.. autoclass:: FormatsNamespace
   :noindex:

When implementing your own formats, you should probably follow the same approach: define a class
and class variables with formats that your code will use. Using FormatsNamespace directly is not
advised as it is subject to change.

For example::

    @ensure_format_uniqueness
    class MyFormats(FormatsNamespace):
        # Empty base as a quick way to create FormatDicts
        base = FormatDict()

        # Fields that combine with base are generally some format traits
        #   and by themselves don't define enough to constitute a format
        my_font_name = base | {'font_name': 'Arial'}
        my_font_size = base | {'font_size': 12}
        my_alignment = base | {'align': 'right'}

        # Fields that combine format traits are intended for use in code
        my_default_font = my_font_name | my_font_size | my_alignment

        # Do keep in mind that FormatsNamespace already has a number of traits predefined
        #   that can be used to create new formats.

Using :func:`ensure_format_uniqueness` will allow you to make sure all of the defined formats
are distinct and do not repeat.