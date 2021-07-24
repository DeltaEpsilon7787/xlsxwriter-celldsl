from io import BytesIO
from unittest.mock import ANY

from _pytest.recwarn import warns
from pytest import fixture, raises
from xlsxwriter import Workbook

from xlsxwriter_celldsl.cell_dsl import ExecutorHelper, cell_dsl_context
from xlsxwriter_celldsl.errors import CellDSLError, ExecutionCellDSLError, MovementCellDSLError
from xlsxwriter_celldsl.formats import FormatsNamespace as F
from xlsxwriter_celldsl.ops import *
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

    def test_rich_format_data_cell_format_shortcut(self):
        assert self.commit([F.default_font, "Alpha", F.default_header, "Beta", F.wrapped]) == [
            chain_rich([
                WriteRich
                    .with_data("Alpha")
                    .with_format(F.default_font),
                WriteRich
                    .with_data("Beta")
                    .with_format(F.default_header)
                    .with_cell_format(F.wrapped)
            ])
        ]


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

    def test_specific_write(self, ws_mock, mocker):
        spy_blank = mocker.spy(ws_mock.ws, 'write_blank')
        spy_formula = mocker.spy(ws_mock.ws, 'write_formula')

        with cell_dsl_context(ws_mock, overwrites_ok=True) as E:
            E.commit([
                Write.with_data_type('blank').with_data('Alpha'),
                Write.with_data_type('formula').with_data('Beta'),
            ])

        spy_blank.assert_any_call(0, 0, 'Alpha', ws_mock.fmt.verify_format(F.default_font))
        spy_formula.assert_any_call(0, 0, 'Beta', ws_mock.fmt.verify_format(F.default_font))

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
                        .with_cell_format(F.wrapped)
                ])
            ])

        spy.assert_called_with(
            0, 1,
            ws_mock.fmt.verify_format(F.default_table_column_font), "Alpha ",
            ws_mock.fmt.verify_format(F.default_font_bold), "Beta!",
            ws_mock.fmt.verify_format(F.default_header), " But gamma...",
            ws_mock.fmt.verify_format(F.default_header), " Yet delta?",
            # with_format implicitly merges with F.default_font
            ws_mock.fmt.verify_format(F.default_font | F.center), " Epsilon!",
            ws_mock.fmt.verify_format(F.wrapped)
        )

    def test_write_rich_degradation(self, ws_mock, mocker):
        spy = mocker.spy(ws_mock.ws, 'write_string')

        with warns(UserWarning):
            with cell_dsl_context(ws_mock) as E:
                E.commit([WriteRich.with_data("Test")])

        spy.assert_any_call(0, 0, "Test", ws_mock.fmt.verify_format(F.default_font))

    def test_movement_fully(self, ws_mock, mocker):
        spy = mocker.spy(ws_mock.ws, 'write')

        with cell_dsl_context(ws_mock, initial_row=10, initial_col=10, overwrites_ok=True) as E:
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
                AtCell.c(100).r(256), Write.with_data("256, 100")
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
                AtCell.r(10).c(10), StackSave,
                AtCell.r(0).c(0),
                DefineNamedRange
                    .with_name("A1_K11")
                    .bottom_right(-1),
                AtCell.r(5).c(5), Save.at("TestRange"),
                AtCell.r(0).c(0),
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

        with cell_dsl_context(ws_mock, overwrites_ok=True) as E:
            E.commit([
                Write.with_data(0),
                Write.with_data(10), 6,
                Write.with_data(20), 6,
                RefArray.top_left(2).bottom_right(0).at("ChartTest"),
                AddBarChart.do([
                    AddBarChart.target.add_series({'values': '=TestSheet!$A$1:$C$1'})
                ]), 6,
                AddLineChart.do([
                    AddLineChart.target.add_series({'values': RefArray.at('ChartTest')})
                ])
            ])

        spy_wb.assert_any_call({"type": "bar", "subtype": None})
        spy_wb.assert_any_call({"type": "line", "subtype": None})

        spy_ws.assert_any_call(0, 2, ANY)
        spy_ws.assert_any_call(0, 3, ANY)

    def test_conditional_format_1(self, ws_mock, mocker):
        spy_ws = mocker.spy(ws_mock.ws, "conditional_format")

        with cell_dsl_context(ws_mock, overwrites_ok=True) as E:
            E.commit([
                Write.with_data(-100), 6,
                Write.with_data(100), 6,
                AddConditionalFormat
                    .top_left(2)
                    .bottom_right(1)
                    .with_options(
                    {
                        'type': 'cell',
                        'criteria': '>=',
                        'value': 0,
                        'format': F.highlight_border
                    })
            ])

        spy_ws.assert_any_call(0, 0, 0, 1, {
            'type': 'cell',
            'criteria': '>=',
            'value': 0,
            'format': ws_mock.fmt.verify_format(F.default_font | F.highlight_border)
        })

    def test_conditional_format_2(self, ws_mock, mocker):
        spy_ws = mocker.spy(ws_mock.ws, "conditional_format")

        with cell_dsl_context(ws_mock) as E:
            E.commit([
                Write.with_data(-100), 6,
                Write.with_data(100), 6,
                AddConditionalFormat
                    .top_left(2)
                    .bottom_right(1)
                    .with_options(
                    {
                        'type': 'cell',
                        'criteria': '>=',
                        'value': 0,
                    })
                    .with_format(F.highlight_border)
            ])

        spy_ws.assert_any_call(0, 0, 0, 1, {
            'type': 'cell',
            'criteria': '>=',
            'value': 0,
            'format': ws_mock.fmt.verify_format(F.default_font | F.highlight_border)
        })

    def test_add_image_1(self, ws_mock, mocker):
        spy_ws = mocker.spy(ws_mock.ws, "insert_image")

        with cell_dsl_context(ws_mock) as E:
            E.commit([
                AddImage
                    .with_filepath("./test.png")
            ])

        spy_ws.assert_any_call(0, 0, './test.png', {})

    def test_add_image_2(self, ws_mock, mocker):
        spy_ws = mocker.spy(ws_mock.ws, "insert_image")

        # 1x1 PNG
        test = BytesIO(bytearray([
            0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52, 0x00,
            0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53, 0xde, 0x00,
            0x00, 0x00, 0x01, 0x73, 0x52, 0x47, 0x42, 0x00, 0xae, 0xce, 0x1c, 0xe9, 0x00, 0x00, 0x00, 0x04, 0x67,
            0x41, 0x4d, 0x41, 0x00, 0x00, 0xb1, 0x8f, 0x0b, 0xfc, 0x61, 0x05, 0x00, 0x00, 0x00, 0x09, 0x70, 0x00,
            0x48, 0x59, 0x73, 0x00, 0x00, 0x0e, 0xc3, 0x00, 0x00, 0x0e, 0xc3, 0x01, 0xc7, 0x6f, 0xa8, 0x64, 0x00,
            0x00, 0x0c, 0x49, 0x44, 0x41, 0x54, 0x18, 0x57, 0x63, 0xf8, 0xff, 0xff, 0x3f, 0x00, 0x05, 0xfe, 0x02,
            0xa7, 0x35, 0x81, 0x84, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae, 0x42, 0x60, 0x82, 0xfe,
        ]))

        with cell_dsl_context(ws_mock) as E:
            E.commit([
                AddImage
                    .with_image_data(test)
            ])

        spy_ws.assert_any_call(0, 0, '', {
            'image_data': test
        })

    def test_add_chart_combined(self, ws_mock, mocker):
        spy_wb = mocker.spy(ws_mock.wb, "add_chart")
        spy_ws = mocker.spy(ws_mock.ws, "insert_chart")

        with cell_dsl_context(ws_mock, overwrites_ok=True) as E:
            E.commit([
                AddBarChart.do([
                    AddBarChart.target.add_series({'values': '=TestSheet!$A$1:$C$1'}),
                    AddBarChart.target.combine(
                        AddLineChart.do([
                            AddLineChart.target.add_series({'values': '=TestSheet!$A$2:$C$2'})
                        ])
                    ),
                ]),
            ])

        spy_wb.assert_any_call({'type': 'bar', 'subtype': None})
        spy_wb.assert_any_call({'type': 'line', 'subtype': None})

        spy_ws.assert_called_once_with(0, 0, ANY)

    def test_set_print_area(self, ws_mock, mocker):
        spy_ws = mocker.spy(ws_mock.ws, "print_area")

        with cell_dsl_context(ws_mock) as E:
            E.commit([
                Save.at('Alpha'), 333,
                Save.at('Beta'),
                SetPrintArea.with_top_left('Alpha').with_bottom_right('Beta')
            ])

        spy_ws.assert_any_call(0, 0, 3, 3)

    def test_write_rich_impose(self, ws_mock, mocker):
        spy = mocker.spy(ws_mock.ws, 'write_rich_string')

        with cell_dsl_context(ws_mock) as E:
            E.commit([
                chain_rich([
                    WriteRich
                        .with_data("Alpha")
                        .with_format(F.default_font),
                    WriteRich
                        .with_data("Beta")
                        .with_format(F.default_font_bold)
                        .with_cell_format(F.wrapped)
                ]),
                ImposeFormat.with_format(F.rotated_90)
            ])

        spy.assert_called_with(
            0, 0,
            ws_mock.fmt.verify_format(F.default_font), "Alpha",
            ws_mock.fmt.verify_format(F.default_font_bold), "Beta",
            ws_mock.fmt.verify_format(F.default_font | F.wrapped | F.rotated_90)
        )

    def test_write_rich_override(self, ws_mock, mocker):
        spy = mocker.spy(ws_mock.ws, 'write_rich_string')

        with cell_dsl_context(ws_mock) as E:
            E.commit([
                chain_rich([
                    WriteRich
                        .with_data("Alpha")
                        .with_format(F.default_font),
                    WriteRich
                        .with_data("Beta")
                        .with_format(F.default_font_bold)
                        .with_cell_format(F.wrapped)
                ]),
                OverrideFormat.with_format(F.rotated_90)
            ])

        spy.assert_called_with(
            0, 0,
            ws_mock.fmt.verify_format(F.default_font), "Alpha",
            ws_mock.fmt.verify_format(F.default_font_bold), "Beta",
            ws_mock.fmt.verify_format(F.default_font | F.rotated_90)
        )

class TestCellDSLErrors:
    def test_negative_coords(self, ws_mock):
        with raises(MovementCellDSLError, match="Illegal coords"):
            with cell_dsl_context(ws_mock) as E:
                E.commit([7])

    def test_beyond_limit_coords_row(self, ws_mock):
        with raises(MovementCellDSLError, match="Illegal coords"):
            with cell_dsl_context(ws_mock) as E:
                E.commit([AtCell.r(10000000000)])

    def test_beyond_limit_coords_col(self, ws_mock):
        with raises(MovementCellDSLError, match="Illegal coords"):
            with cell_dsl_context(ws_mock) as E:
                E.commit([AtCell.c(10000000000)])

    def test_nonexistent_save_point(self, ws_mock):
        with raises(MovementCellDSLError, match="Save point TYPO SAVE"):
            with cell_dsl_context(ws_mock) as E:
                E.commit([Save.at("SAVE POINT 1"), 6, Save.at("SAVE POINT 2"), 6, Load.at("TYPO SAVE")])

    def test_backtrack_too_far(self, ws_mock):
        with raises(MovementCellDSLError, match="backtrack 100"):
            with cell_dsl_context(ws_mock) as E:
                E.commit([Write.with_data("alpha"), 6, Write.with_data("beta"), 6, BacktrackCell.rewind(100)])

    def test_load_from_empty_save_stack(self, ws_mock):
        with raises(MovementCellDSLError, match="is empty"):
            with cell_dsl_context(ws_mock) as E:
                E.commit([StackLoad])

    def test_range_fail_11(self, ws_mock):
        with raises(MovementCellDSLError, match="Top left corner would use 100"):
            with cell_dsl_context(ws_mock) as E:
                E.commit([DrawBoxBorder.top_left(100)])

    def test_range_fail_12(self, ws_mock):
        with raises(MovementCellDSLError, match="Top left corner would look 100"):
            with cell_dsl_context(ws_mock) as E:
                E.commit([DrawBoxBorder.top_left(-100)])

    def test_range_fail_21(self, ws_mock):
        with raises(MovementCellDSLError, match="Bottom right corner would use 100"):
            with cell_dsl_context(ws_mock) as E:
                E.commit([DrawBoxBorder.bottom_right(100)])

    def test_range_fail_22(self, ws_mock):
        with raises(MovementCellDSLError, match="Bottom right corner would look 100"):
            with cell_dsl_context(ws_mock) as E:
                E.commit([DrawBoxBorder.bottom_right(-100)])

    def test_range_fail_31(self, ws_mock):
        with raises(CellDSLError, match="Tried to use a save point named FAIL for top left"):
            with cell_dsl_context(ws_mock) as E:
                E.commit([DrawBoxBorder.top_left("FAIL")])

    def test_range_fail_32(self, ws_mock):
        with raises(CellDSLError, match="Tried to use a save point named FAIL for bottom right"):
            with cell_dsl_context(ws_mock) as E:
                E.commit([DrawBoxBorder.bottom_right("FAIL")])

    def test_double_overwrite_format(self, ws_mock):
        with raises(CellDSLError, match="There's already an OverrideFormat for cell"):
            with cell_dsl_context(ws_mock) as E:
                E.commit([OverrideFormat.with_format(F.default_font), OverrideFormat.with_format(F.default_font_bold)])

    def test_name_stack_tracking_basic(self, ws_mock):
        with raises(MovementCellDSLError, match='Illegal coords') as exc:
            with cell_dsl_context(ws_mock) as E:
                E.commit([SectionBegin.with_name("Section1"), 7])

        assert "Name stack: ['Section1']" in str(exc.value)

    def test_name_stack_tracking_advanced(self, ws_mock):
        with raises(MovementCellDSLError, match='Illegal coords') as exc:
            with cell_dsl_context(ws_mock) as E:
                E.commit([
                    SectionBegin.with_name("Section1"), [
                        SectionBegin.with_name("Section2"), [
                            SectionBegin.with_name("Section3"),
                            SectionEnd,
                        ],
                        7
                    ]
                ])

        assert "Name stack: ['Section2', 'Section1']" in str(exc.value)

    def test_non_empty_name_stack(self, ws_mock):
        with raises(MovementCellDSLError, match='Name stack is not empty') as exc:
            with cell_dsl_context(ws_mock) as E:
                E.commit([
                    SectionBegin.with_name("Section1"), [
                        SectionBegin.with_name("Section2"),
                        66,
                        SectionEnd
                    ]
                ])

        assert "Name stack: ['Section1']" in str(exc.value)

    def test_overwrite_protection_same_action(self, ws_mock):
        with cell_dsl_context(ws_mock) as E:
            E.commit([
                "0, 0", 6,
                F.default_font, "0, 1", 6,
                "0, 2", 4,
                F.default_font, "0, 1"
            ])

        with raises(ExecutionCellDSLError, match=r'Overwrite has occurred at \(0, 1\)'):
            with cell_dsl_context(ws_mock) as E:
                E.commit([
                    "0, 0", 6,
                    F.default_font_centered, "0, 1", 6,
                    "0, 2", 4,
                    F.default_font, "0, 1"
                ])

    def test_conditional_format_ambiguity(self, ws_mock):
        with raises(CellDSLError, match=r'Both format key and format field are specified, use only one of them.'):
            with cell_dsl_context(ws_mock) as E:
                E.commit([
                    AddConditionalFormat.with_format(F.default_font).with_options({
                        'format': F.highlight_border
                    })
                ])

    def test_name_stack_coord_unbound(self, ws_mock):
        with raises(MovementCellDSLError, match='Illegal coords') as exc:
            with cell_dsl_context(ws_mock) as E:
                E.commit([
                    66, SectionBegin.with_name('Section1'), [
                        4, SectionBegin.with_name('Section2'), [
                            7,
                            SectionEnd,
                        ],
                        SectionEnd,
                    ],
                ])

        assert "Name stack: ['Section2', 'Section1']" in str(exc.value)
