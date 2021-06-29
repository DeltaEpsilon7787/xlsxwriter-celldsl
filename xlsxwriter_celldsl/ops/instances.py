from typing import TYPE_CHECKING

from ops import classes

if TYPE_CHECKING:
    from xlsxwriter.chart_area import ChartArea
    from xlsxwriter.chart_bar import ChartBar
    from xlsxwriter.chart_column import ChartColumn
    from xlsxwriter.chart_doughnut import ChartDoughnut
    from xlsxwriter.chart_line import ChartLine
    from xlsxwriter.chart_pie import ChartPie
    from xlsxwriter.chart_radar import ChartRadar
    from xlsxwriter.chart_scatter import ChartScatter
    from xlsxwriter.chart_stock import ChartStock

StackSave = classes.StackSaveOp()
StackLoad = classes.StackLoadOp()
Load = classes.LoadOp()
Save = classes.SaveOp()
RefArray = classes.RefArrayOp()
SectionBegin = classes.SectionBeginOp()
SectionEnd = classes.SectionEndOp()

Move = classes.MoveOp()
AtCell = classes.AtCellOp()
BacktrackCell = classes.BacktrackCellOp()

Write = classes.WriteOp()
WriteNumber = Write.with_data_type('number')
WriteString = Write.with_data_type('string')
WriteBlank = Write.with_data_type('blank')
WriteFormula = Write.with_data_type('formula')
WriteDatetime = Write.with_data_type('datetime')
WriteBoolean = Write.with_data_type('boolean')
WriteURL = Write.with_data_type('url')

MergeWrite = classes.MergeWriteOp()
WriteRich = classes.WriteRichOp()

ImposeFormat = classes.ImposeFormatOp()
OverrideFormat = classes.OverrideFormatOp()

DrawBoxBorder = classes.DrawBoxBorderOp()
DefineNamedRange = classes.DefineNamedRangeOp()

SetRowHeight = classes.SetRowHeightOp()
SetColWidth = classes.SetColumnWidthOp()

SubmitHPagebreak = classes.SubmitHPagebreakOp()
SubmitVPagebreak = classes.SubmitVPagebreakOp()
ApplyPagebreaks = classes.ApplyPagebreaksOp()

NextRow = Move.r(1)
NextCol = Move.c(1)
PrevRow = Move.r(-1)
PrevCol = Move.c(-1)
NextRowSkip = Move.r(2)
NextColSkip = Move.c(2)
PrevRowSkip = Move.r(-2)
PrevColSkip = Move.c(-2)

AddComment = classes.AddCommentOp()

AddAreaChart: classes.AddChartOp['ChartArea'] = classes.AddChartOp(type='area')
AddBarChart: classes.AddChartOp['ChartBar'] = classes.AddChartOp(type='bar')
AddColumnChart: classes.AddChartOp['ChartColumn'] = classes.AddChartOp(type='column')
AddLineChart: classes.AddChartOp['ChartLine'] = classes.AddChartOp(type='line')
AddPieChart: classes.AddChartOp['ChartPie'] = classes.AddChartOp(type='pie')
AddDoughnutChart: classes.AddChartOp['ChartDoughnut'] = classes.AddChartOp(type='doughnut')
AddScatterChart: classes.AddChartOp['ChartScatter'] = classes.AddChartOp(type='scatter')
AddStockChart: classes.AddChartOp['ChartStock'] = classes.AddChartOp(type='stock')
AddRadarChart: classes.AddChartOp['ChartRadar'] = classes.AddChartOp(type='radar')

AddConditionalFormat = classes.AddConditionalFormatOp()

AddImage = classes.AddImageOp()
