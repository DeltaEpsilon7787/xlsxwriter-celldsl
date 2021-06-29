# XlsxWriter-CellDSL

## What is this?

**XlsxWriter-CellDSL** is a complementary Python module for the
excellent [XlsxWriter](https://github.com/jmcnamara/XlsxWriter) library that provides a DSL (domain specific language)
to perform common operations in worksheets without having to specify absolute coordinates and keep track of them,
instead opting to use primarily relative movement of an imaginary "cursor" within the spreadsheet, among other things.

## The issue with absolute coordinates

If you've ever written code that generates structures that have a dynamic layout in Excel, you may have noticed that, in
order to make sure writes happen in correct cells, you have to carry data that have been used to figure out the size of
those structures and then sum up those size data with some offsets in order to get the coordinate that you then pass
into `write` function.

This is really painful to do in more complex cases where many structures are present within a worksheet since you have
to keep track of the sizes of each structure and drag this information across every module that's related to the
worksheet.

If you want to then refactor the code or move structures around, you'll have to rewrite the coordinate calculations for
every structure downstream, which is extremely error-prone. Moreover, if you then refactor writing some structures into
functions, you will have to pass into the function some kind of _initial_ coordinates, which is information that's
unlikely to be relevant for the structure itself. To put it simply, a structure doesn't care where it is located and
thus, keeping track of absolute coordinates is a very error-prone and redundant activity.

To put it another way, in many cases writing a structure into an Excel sheet is not a random access operation for which
absolute coordinates would be more appropriate for. Writes do not occur to arbitrary cells, in fact, more often than not
writes occur in some kind of sequential order, with known spacing and local positioning, but not necessarily known
global positioning.

Not only that, but even writing **and** reading an Excel sheet follows this principle: at any given moment the operator
cares more about nearby cells than distant ones. Mirroring this mode of operation is what relative coordinates are best
suited for. This module implements a number of utilities that allow the developer to have an imaginary "cursor"
placed somewhere in the worksheet and operations occur wherever this cursor is, followed by moving the cursor with arrow
keys into the next position.

## Features and uses

* `MoveOp`: Move the cursor around using relative coordinates.
* `AtCellOp`: Perform an absolute coordinate jump if no other movement option suffices.
* `FormatDict`: Construct formats by treating them as a composition of smaller formats instead of raw dictionaries with
  repeated key-value pairs.
* `FormatHandler`: Delegate keeping track of added formats to XlsxWriter-CellDSL and remove the need to distribute
  references to added formats between generating functions.
* `SaveOp`, `LoadOp`: Give current position a name and then jump back to it later or use it to retrieve the absolute
  coordinates of some point of interest after the script has finished execution.
* `StackSaveOp`, `StackLoadOp`: A structure composed of substructures would want to take advantage of `SaveOp`
  capabilities without having to generate a name for it.
* `WriteOp`, `MergeWriteOp`, `WriteRichOp`: Perform common writing actions to current cell, only focusing on data, and
  the cell format.
* `ImposeFormatOp`, `OverrideFormatOp`, `DrawBoxBorderOp`: Deferred execution of operations allows additional formatting
  to be applied to writing actions after they occur which would ordinarily require changing the arguments of the first
  writing function call.
* `SectionBeginOp`, `SectionEndOp`: Errors are inevitable and though deferred execution makes debugging more difficult,
  this needn't be the case if you annotate segments with names.
* `AddChartOp`: Add charts to the worksheet and utilize the flexibility of named ranges in an environment which does not
  allow named ranges with `RefArrayOp`.
* Exceptions provide a lot of useful information to track down the line that causes it.
* Several short forms of common operations improve conciseness of code.
* Deferred execution of operations allows taking advantage of `constant_memory` mode in XlsxWriter easily, without
  having to contend with write-to-stream limitations such as ensuring the writes occurs in left-to-right, top-to-bottom
  order only, thus providing the performance of `constant_memory` mode, but flexibility of regular mode.
* Deferred execution allows introspection into the action chain and modifying it out-of-order.
* Upon execution, history of operations can be saved and used in scripts further down.
* Avoid more bugs by preventing overwriting data over non-emtpy cells with `overwrites_ok` attribute.

## Documentation

Read the full documentation [here](https://xlsxwriter-celldsl.readthedocs.io/en/latest/).

## Usage example

```py
from xlsxwriter import Workbook

# Various operations
import xlsxwriter_celldsl.ops as ops
# The entry point to the library
from xlsxwriter_celldsl import cell_dsl_context
# A factory of objects needed for the context manager
from xlsxwriter_celldsl.utils import WorkbookPair
# A number of basic formats
from xlsxwriter_celldsl.formats import FormatsNamespace as F
# Useful functions to assist in printing sequences
from xlsxwriter_celldsl.utils import row_chain, col_chain, segment

wb = Workbook('out.xlsx')
wb_pair = WorkbookPair.from_wb(wb)
ws_triplet = wb_pair.add_worksheet("TestSheet1")

with cell_dsl_context(ws_triplet) as E:
  # ExecutorHelper (as E here) is a special preprocessor object that keeps track of operations
  # to be done and performs some preprocessing on them
  # See the docs for `ExecutorHelper.commit`
  E.commit([
    # xlsxwriter_celldsl.ops exports both command classes and basic instances of those classes.
    # ExecutorHelper.commit uses instances.
    # All commands are immutable objects, however, they are cached and reused
    #   so few new instances are created.
    ops.Write
      .with_data("Hello, world, at A1, using left aligned Liberation Sans 10 (default font)!")
      .with_format(F.default_font),
    ops.Move.c(3),  # Move three columns to the right
    "Wow, short form of ops.Write.with_data('this string'), at D1, three columns away from A1!",
    11,  # Short form of ops.Move, refer to ExecutorHelper.commit to see how this works
    F.default_font_bold, "Wow, I'm at B3 now, written in bold", 2,
    [
      [
        [
          "However deeply I'm nested, I will be reached anyway, at B4"
        ]
      ]
    ], 6,
    # Rich string short form, several formats within a single text cell
    F.default_font, "A single cell, but two parts, first half normal ",
    F.default_font_bold, "but second half bold! For as long as we stay at C4...", 6,
    "Oops, D4 now",
    # Saving current position as "see you later"
    ops.Save.at("see you later"),
    # Absolute coordinate jump
    ops.AtCell.r(49).c(1), "Jumping all the way to B50",
    # Jumping to some previously saved position
    ops.Load.at("see you later"), 6, "We've gone back to D4, moved right and now it's E4",
    3333,
    ops.Save.at("Bottom Right Corner"),
    # Reversing movement back in time
    ops.BacktrackCell.rewind(1),
    # Drawing a box using borders
    ops.DrawBoxBorder.bottom_right("Bottom Right Corner"), 33,
    # Two formats may be "merged" together using OR operator
    # In this case, we add "wrapped" trait to default font
    F.default_font | F.wrapped, "And now, we're inside a 5x5 box, starting at E4, but this is G6."
                                "Even though this operation precedes the next one, the next one affect this cell"
                                ", thus we are inside a smaller box that only encloses G6.",
    ops.DrawBoxBorder,
    ops.AtCell.r(10).c(0),
    # Sections allow you to document your code segments by giving them names
    #   and also assist in debugging as you will be shown the name stack
    #     up until the line that causes the exception
    ops.SectionBegin.with_name("Multiplication table"), [
      # col_chain / row_chain write data sequentially from an iterable
      # row_chain prints it in a row, but the actual position of the cursor doesn't change!            
      "A sequence from 1 to 9, horizontally", 6, row_chain([
        f"* {v}"
        for v in range(1, 10)
      ]), 1,
      # col_chain prints it in a column
      "A sequence from 1 to 9, vertically", 2, col_chain([
        f"{v} *"
        for v in range(1, 10)
      ]), 6,
      # Nothing stops you from chaining chains
      col_chain([
        row_chain([
          ops.Write.with_data(a * b)
          for b in range(1, 10)
        ])
        for a in range(1, 10)
      ]),
      # Every SectionBegin must be matched with a SectionEnd
      ops.SectionEnd,
      # ...however you can skip that by using utils.segment to implicitly add SectionBegin and SectionEnd to
      #   a piece of code
      segment("Empty segment", [])
    ]
  ])
```

# Changelog

## 0.5.0

* Charts can now be combined.
* A suite of WriteOp variations for writing data with known types like `WriteNumber` and `WriteDatetime`.
* **BREAKING CHANGE**: Operation instances and classes have been separated: the classes are now in `ops.classes` module.
  Regular instances are still importable from `ops`, so unless your code creates its own instances of operations, you
  needn't change anything.

## 0.4.0

* Add `AddConditionalFormatOp` and `AddImageOp`
* Write the format section in the docs
* New default format traits in `FormatsNamespace`

## 0.3.0

* Add `overwrites_ok`
* Docs!
* Removed dummy_cell_dsl_context
* Complete overhaul to `AddChartOp`, removing the string function name interface

## 0.2.0

* Add `SectionBeginOp` and `SectionEndOp`
* Improvement to error reporting: now they provide some context
* Remove format data from repr of commands
* Separate `CellDSLError` into `MovementCellDSLError` and `ExecutionCellDSLError`
* Raise exceptions on various error that may occur from XlsxWriter side (use proper exceptions instead of return codes)