from io import BytesIO
from unittest.mock import ANY

from pytest import fixture, raises
from xlsxwriter import Workbook
from xlsxwriter.chart_bar import ChartBar
from xlsxwriter.chart_line import ChartLine

from xlsxwriter_celldsl import FormatsNamespace as F
from xlsxwriter_celldsl.cell_dsl import CellDSLError, ExecutorHelper, cell_dsl_context
from xlsxwriter_celldsl.ops import AddBarChart, AddLineChart, AtCell, BacktrackCell, DefineNamedRange, ImposeFormat, \
    Load, MergeWrite, \
    Move, \
    OverrideFormat, RefArray, Save, \
    StackLoad, \
    StackSave, Write, \
    WriteRich
from xlsxwriter_celldsl.utils import WorkbookPair, chain_rich


class TestExecutorHelper:
    @staticmethod
    def commit(lst):
        e = ExecutorHelper()
        e.commit(lst)

        return e.action_chain

    def test_movement_preprocessor(self):
        assert self.commit([1]) == [Move.r(1).c(-1)]
        assert self.commit([2]) == [Move.r(1).c(0)]
        assert self.commit([3]) == [Move.r(1).c(1)]
        assert self.commit([4]) == [Move.r(0).c(-1)]
        assert self.commit([5]) == [Move.r(0).c(0)]
        assert self.commit([6]) == [Move.r(0).c(1)]
        assert self.commit([7]) == [Move.r(-1).c(-1)]
        assert self.commit([8]) == [Move.r(-1).c(0)]
        assert self.commit([9]) == [Move.r(-1).c(1)]

    def test_movement_simplification(self):
        assert self.commit([91734682]) == [Move]

    def test_nested_flattening(self):
        assert self.commit([5, [5]]) == [Move, Move]
        assert self.commit([5, [5, [5]], 5]) == [Move, Move, Move, Move]

    def test_data_write_shortcut(self):
        assert self.commit(["Test"]) == [Write.with_data("Test")]

    def test_format_data_write_shortcut(self):
        assert self.commit([F.default_font, "Test"]) == [
            Write
                .with_data("Test")
                .with_format(F.default_font)
        ]

    def test_format_merge_data_write_shortcut(self):
        assert self.commit([F.default_font, F.center, "Test"]) == [
            Write
                .with_data("Test")
                .with_format(F.default_font | F.center)
        ]

    def test_rich_format_data_shortcut(self):
        assert self.commit([F.default_font, "Alpha", F.default_header, "Beta"]) == [
            chain_rich([
                WriteRich
                    .with_data("Alpha")
                    .with_format(F.default_font),
                WriteRich
                    .with_data("Beta")
                    .with_format(F.default_header)
            ])
        ]

    def test_format_nonformat_error(self):
        with raises(CellDSLError):
            self.commit([F.default_font, None])


@fixture
def ws_mock():
    dump = BytesIO()

    wb = Workbook(dump, {'constant_memory': True})
    pair = WorkbookPair.from_wb(wb)
    result = pair.add_worksheet("TestSheet")

    yield result

    wb.close()


class TestCellDSL:
    def test_write(self, ws_mock, mocker):
        spy = mocker.spy(ws_mock.ws, 'write')

        with cell_dsl_context(ws_mock) as E:
            E.commit(["Alpha", 3, F.default_font_centered, "Beta"])

        spy.assert_any_call(0, 0, 'Alpha', ws_mock.fmt.verify_format(F.default_font))
        spy.assert_any_call(1, 1, 'Beta', ws_mock.fmt.verify_format(F.default_font_centered))

    def test_write_fail(self, ws_mock):
        with raises(CellDSLError):
            with cell_dsl_context(ws_mock) as E:
                E.commit([7, "Out of bounds write"])

    def test_merge_write(self, ws_mock, mocker):
        spy = mocker.spy(ws_mock.ws, 'merge_range')

        with cell_dsl_context(ws_mock) as E:
            E.commit([
                6,
                MergeWrite
                    .with_data("Alpha")
                    .with_format(F.default_percent)
                    .with_size(10)
            ])

        spy.assert_called_with(0, 1, 0, 11, "Alpha", ws_mock.fmt.verify_format(F.default_percent))

    def test_write_rich(self, ws_mock, mocker):
        spy = mocker.spy(ws_mock.ws, 'write_rich_string')

        with cell_dsl_context(ws_mock) as E:
            E.commit([
                6,
                chain_rich([
                    WriteRich
                        .with_default_format(F.default_header)
                        .with_data("Alpha ")
                        .with_format(F.default_table_column_font),
                    WriteRich
                        .with_data("Beta!")
                        .with_format(F.default_font_bold),
                    WriteRich
                        .with_data(" But gamma..."),
                    WriteRich
                        .with_data(" Yet delta?"),
                    WriteRich
                        .with_data(" Epsilon!")
                        .with_format(F.center)
                ])
            ])

        spy.assert_called_with(
            0, 1,
            ws_mock.fmt.verify_format(F.default_table_column_font), "Alpha ",
            ws_mock.fmt.verify_format(F.default_font_bold), "Beta!",
            ws_mock.fmt.verify_format(F.default_header), " But gamma...",
            ws_mock.fmt.verify_format(F.default_header), " Yet delta?",
            # with_format implicitly merges with F.default_font
            ws_mock.fmt.verify_format(F.default_font | F.center), " Epsilon!"
        )

    def test_movement_fully(self, ws_mock, mocker):
        spy = mocker.spy(ws_mock.ws, 'write')

        with cell_dsl_context(ws_mock, initial_row=10, initial_col=10) as E:
            E.commit([
                # Basic movement
                Write.with_data("10, 10"), 6,
                Write.with_data("10, 11"), 44,
                Write.with_data("10, 9"), 8,
                Write.with_data("9, 9"),
                # Stack saves
                StackSave, 6666, Write.with_data("9, 13"), 6666,
                StackLoad, Write.with_data("9, 9"),
                # Save-load
                Save.at("Alpha"), 6,
                Save.at("Beta"), 44444, Write.with_data("9, 5"),
                Load.at("Alpha"), Write.with_data("9, 9"),
                Load.at("Beta"), Write.with_data("9, 10"),
                # Backtracking
                BacktrackCell.rewind(2), Write.with_data("9, 5"),
                # Go-to
                AtCell.at_col(100).at_row(256), Write.with_data("256, 100")
            ])

        fmt = ws_mock.fmt.verify_format(F.default_font)

        spy.assert_any_call(9, 5, "9, 5", fmt)
        spy.assert_any_call(9, 5, "9, 5", fmt)
        spy.assert_any_call(9, 9, "9, 9", fmt)
        spy.assert_any_call(9, 9, "9, 9", fmt)
        spy.assert_any_call(9, 9, "9, 9", fmt)
        spy.assert_any_call(9, 10, "9, 10", fmt)
        spy.assert_any_call(9, 13, "9, 13", fmt)
        spy.assert_any_call(10, 9, "10, 9", fmt)
        spy.assert_any_call(10, 10, "10, 10", fmt)
        spy.assert_any_call(10, 11, "10, 11", fmt)
        spy.assert_any_call(256, 100, "256, 100", fmt)

    def test_range_trait(self, ws_mock, mocker):
        spy = mocker.spy(ws_mock.wb, 'define_name')

        with cell_dsl_context(ws_mock) as E:
            E.commit([
                666,
                DefineNamedRange
                    .with_name("A1_D1")
                    .top_left(1),
                AtCell.at_row(10).at_col(10), StackSave,
                AtCell.at_row(0).at_col(0),
                DefineNamedRange
                    .with_name("A1_K11")
                    .bottom_right(-1),
                AtCell.at_row(5).at_col(5), Save.at("TestRange"),
                AtCell.at_row(0).at_col(0),
                DefineNamedRange
                    .with_name("A1_F6")
                    .bottom_right("TestRange"),
            ])

        spy.assert_any_call("A1_D1", "=TestSheet!$A$1:$D$1")
        spy.assert_any_call("A1_F6", "=TestSheet!$A$1:$F$6")
        spy.assert_any_call("A1_K11", "=TestSheet!$A$1:$K$11")

    def test_impose_format(self, ws_mock, mocker):
        spy = mocker.spy(ws_mock.ws, 'write')

        with cell_dsl_context(ws_mock) as E:
            E.commit([
                Write.with_data("Test").with_format(F.default_font_bold),
                ImposeFormat.with_format(F.center)
            ])

        spy.assert_called_with(0, 0, "Test", ws_mock.fmt.verify_format(F.default_font_bold | F.center))

    def test_override_format(self, ws_mock, mocker):
        spy = mocker.spy(ws_mock.ws, 'write')

        with cell_dsl_context(ws_mock) as E:
            E.commit([
                Write.with_data("Test").with_format(F.default_font_bold),
                OverrideFormat.with_format(F.default_font_centered)
            ])

        spy.assert_called_with(0, 0, "Test", ws_mock.fmt.verify_format(F.default_font_centered))

    def test_add_chart(self, ws_mock, mocker):
        spy_wb = mocker.spy(ws_mock.wb, "add_chart")
        spy_ws = mocker.spy(ws_mock.ws, "insert_chart")

        with cell_dsl_context(ws_mock) as E:
            E.commit([
                Write.with_data(0),
                Write.with_data(10), 6,
                Write.with_data(20), 6,
                RefArray.top_left(2).bottom_right(0).at("ChartTest"),
                AddBarChart
                    .do(ChartBar.add_series, ({'values': '=TestSheet!$A$1:$C$1'})), 6,
                AddLineChart
                    .do(ChartLine.add_series, ({'values': RefArray.at("ChartTest")})),
            ])

        spy_wb.assert_any_call({"type": "bar", "subtype": None})
        spy_wb.assert_any_call({"type": "line", "subtype": None})

        spy_ws.assert_any_call(0, 2, ANY)
        spy_ws.assert_any_call(0, 3, ANY)
