using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Text.Json;
using Microsoft.Office.Interop.Excel;

namespace AgentForExcel.Operations.Dashboard
{
    /// <summary>基于同一透视缓存创建可由切片器联动的原生 Excel 数据看板。</summary>
    public sealed class CreateDashboardOp : IOperation
    {
        public string ToolName => "dashboard_create";
        public bool IsWriteOperation => true;

        private readonly string _sourceSheet;
        private readonly string _sourceAddress;
        private readonly string _dashboardSheetName;
        private readonly string _title;
        private readonly string _dateField;
        private readonly string _categoryField;
        private readonly string _seriesField;
        private readonly string _valueField;
        private readonly string[] _filterFields;
        private readonly string _filterMode;
        private readonly string _aggregation;
        private readonly int _topN;
        private readonly string _numberFormat;

        private CreateDashboardOp(
            string sourceSheet, string sourceAddress, string dashboardSheetName, string title,
            string dateField, string categoryField, string seriesField, string valueField,
            string[] filterFields, string filterMode, string aggregation, int topN, string numberFormat)
        {
            _sourceSheet = sourceSheet;
            _sourceAddress = sourceAddress;
            _dashboardSheetName = dashboardSheetName;
            _title = title;
            _dateField = dateField;
            _categoryField = categoryField;
            _seriesField = seriesField;
            _valueField = valueField;
            _filterFields = filterFields ?? new string[0];
            _filterMode = filterMode;
            _aggregation = aggregation;
            _topN = topN;
            _numberFormat = numberFormat;
        }

        public string Describe()
        {
            var source = (string.IsNullOrWhiteSpace(_sourceSheet) ? "活动工作表" : "工作表「" + _sourceSheet + "」") + "!" + _sourceAddress;
            var mode = _filterMode == "dropdown" ? "下拉兼容筛选" : "原生切片器";
            return $"根据 {source} 创建联动数据看板「{_title}」，包含 KPI、趋势、排名、占比、对比、明细透视表和 {_filterFields.Length} 个全局筛选器（{mode}）；源数据不会被修改";
        }

        public string Execute(AppContext context)
        {
            // 看板构建含大量工作表写入与多个透视/图表对象创建,批量作用域抑制逐次重绘。
            using (new ExcelBatchScope(context))
            {
                return ExecuteCore(context);
            }
        }

        private string ExecuteCore(AppContext context)
        {
            var workbook = context?.Excel?.ActiveWorkbook;
            if (workbook == null) throw new InvalidOperationException("当前没有打开的工作簿。");

            var sourceSheet = Cell.CellOperationSupport.GetWorksheet(context, _sourceSheet);
            var sourceRange = Cell.CellOperationSupport.GetRange(sourceSheet, _sourceAddress);
            if (sourceRange.Rows.Count < 2 || sourceRange.Columns.Count < 2)
                throw new ArgumentException("看板源区域至少需要 2 行 × 2 列，并包含标题行。");

            ValidateSource(sourceRange);

            Worksheet dashboardSheet = null;
            Worksheet supportSheet = null;
            var slicerCaches = new List<dynamic>();
            string dropdownBindingName = null;
            try
            {
                dashboardSheet = Analysis.AnalysisSheetSupport.CreateUniqueWorksheet(context, _dashboardSheetName);
                supportSheet = Analysis.AnalysisSheetSupport.CreateUniqueWorksheet(context, _dashboardSheetName + "数据");
                BuildDashboardCanvas(dashboardSheet, sourceSheet, sourceRange);

                var caches = workbook.PivotCaches();
                var sourceReference = sourceRange.get_Address(
                    true, true, XlReferenceStyle.xlR1C1, true, Type.Missing);
                var cache = caches.Create(XlPivotTableSourceType.xlDatabase, sourceReference, Type.Missing);
                cache.MissingItemsLimit = XlPivotTableMissingItems.xlMissingItemsNone;

                var pivots = new List<PivotTable>();
                var totalPivot = BuildMetricPivot(cache, supportSheet, "A1", "AgentKpiTotal", "合计", XlConsolidationFunction.xlSum, _numberFormat);
                var averagePivot = BuildMetricPivot(cache, supportSheet, "D1", "AgentKpiAverage", "平均值", XlConsolidationFunction.xlAverage, _numberFormat);
                var countPivot = BuildMetricPivot(cache, supportSheet, "G1", "AgentKpiCount", "记录数", XlConsolidationFunction.xlCount, "#,##0");
                pivots.Add(totalPivot);
                pivots.Add(averagePivot);
                pivots.Add(countPivot);

                var rankingPivot = BuildBreakdownPivot(
                    cache, supportSheet, "A20", "AgentRanking", _categoryField,
                    "指标值", ParseAggregation(_aggregation), true, _topN);
                pivots.Add(rankingPivot);

                var trendDimension = !string.IsNullOrWhiteSpace(_dateField)
                    ? _dateField
                    : (!string.IsNullOrWhiteSpace(_seriesField) ? _seriesField : _categoryField);
                var trendPivot = BuildBreakdownPivot(
                    cache, supportSheet, "J20", "AgentTrend", trendDimension,
                    "指标值", ParseAggregation(_aggregation), false, 0);
                pivots.Add(trendPivot);

                var shareDimension = !string.IsNullOrWhiteSpace(_seriesField) ? _seriesField : _categoryField;
                var sharePivot = BuildBreakdownPivot(
                    cache, supportSheet, "S20", "AgentShare", shareDimension,
                    "指标值", ParseAggregation(_aggregation), true, 0);
                pivots.Add(sharePivot);
                var compareDimension = !string.IsNullOrWhiteSpace(_seriesField) ? _seriesField : _categoryField;
                var comparePivot = BuildBreakdownPivot(
                    cache, supportSheet, "AC20", "AgentCompare", compareDimension,
                    "指标值", ParseAggregation(_aggregation), true, 0);
                pivots.Add(comparePivot);

                var detailPivot = BuildDetailPivot(cache, dashboardSheet, "A64", "AgentDetail");
                pivots.Add(detailPivot);

                var actualFilterMode = _filterFields.Length == 0 ? "none" : _filterMode;
                if (_filterFields.Length > 0 && _filterMode != "dropdown")
                {
                    try
                    {
                        for (var i = 0; i < _filterFields.Length; i++)
                        {
                            var slicerCache = CreateSlicer(workbook, dashboardSheet, pivots, _filterFields[i], i);
                            slicerCaches.Add(slicerCache);
                        }
                        ConnectAllSlicerCaches(slicerCaches, pivots);
                        actualFilterMode = "slicer";
                    }
                    catch when (_filterMode == "auto")
                    {
                        for (var i = slicerCaches.Count - 1; i >= 0; i--)
                            Try(() => slicerCaches[i].Delete());
                        slicerCaches.Clear();
                        actualFilterMode = "dropdown";
                    }
                }

                if (_filterFields.Length > 0 && actualFilterMode == "dropdown")
                {
                    PrepareDropdownPivotFields(pivots, _filterFields);
                    dropdownBindingName = CreateDropdownFilters(
                        context.Excel, workbook, dashboardSheet, supportSheet, sourceRange, pivots, _filterFields);
                }

                BindMetricCard(dashboardSheet, "A11:D14", totalPivot, _numberFormat);
                BindMetricCard(dashboardSheet, "E11:H14", averagePivot, _numberFormat);
                BindMetricCard(dashboardSheet, "I11:L14", countPivot, "#,##0");

                CreateChart(dashboardSheet, rankingPivot, "A16", "G34", XlChartType.xlBarClustered,
                    _categoryField + "排名 Top " + _topN, ChartStyle.Ranking);
                CreateChart(dashboardSheet, trendPivot, "H16", "N34", XlChartType.xlLineMarkers,
                    (!string.IsNullOrWhiteSpace(_dateField) ? _dateField + "趋势" : trendDimension + "分布"), ChartStyle.Trend);
                CreateChart(dashboardSheet, sharePivot, "A36", "G54", XlChartType.xlDoughnut,
                    shareDimension + "占比", ChartStyle.Share);
                CreateChart(dashboardSheet, comparePivot, "H36", "N54", XlChartType.xlColumnClustered,
                    compareDimension + "对比", ChartStyle.Comparison);

                if (actualFilterMode == "dropdown")
                    DashboardInteractionManager.RefreshDashboard(workbook, dashboardSheet.Name);

                supportSheet.Visible = XlSheetVisibility.xlSheetVeryHidden;
                try
                {
                    ThisAddIn.Log("看板构建完成: ChartObjects.Count=" +
                        Convert.ToInt32(((ChartObjects)dashboardSheet.ChartObjects(Type.Missing)).Count) +
                        ", FilterMode=" + actualFilterMode);
                }
                catch { }
                dashboardSheet.Activate();
                try { context.Excel.ActiveWindow.DisplayGridlines = false; } catch { }
                try { context.Excel.ActiveWindow.Zoom = 85; } catch { }

                var modeText = actualFilterMode == "slicer" ? "原生切片器" :
                    actualFilterMode == "dropdown" ? "公式卡片＋下拉兼容筛选" : "无筛选器";
                return $"已创建联动数据看板「{dashboardSheet.Name}」，包含 3 个 KPI、4 个动态透视图、{_filterFields.Length} 个全局筛选器和 1 个明细透视表；当前使用{modeText}，源工作表未修改。";
            }
            catch
            {
                if (!string.IsNullOrWhiteSpace(dropdownBindingName))
                    DashboardInteractionManager.DeleteDefinition(workbook, dropdownBindingName);
                for (var i = slicerCaches.Count - 1; i >= 0; i--)
                    Try(() => slicerCaches[i].Delete());
                if (supportSheet != null)
                    Analysis.AnalysisSheetSupport.DeleteWorksheetSilently(context, supportSheet);
                if (dashboardSheet != null)
                    Analysis.AnalysisSheetSupport.DeleteWorksheetSilently(context, dashboardSheet);
                throw;
            }
        }

        private void ValidateSource(Range sourceRange)
        {
            // 一次读入整个区域,表头校验与数值探测全部走托管数组,
            // 避免对上万行源数据逐格发起 COM 调用。
            var values = sourceRange.Value2 as object[,];
            if (values == null)
                throw new ArgumentException("看板源区域数据无法读取。");
            var rowCount = values.GetLength(0);
            var columnCount = values.GetLength(1);

            var headers = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (var column = 1; column <= columnCount; column++)
            {
                var header = Convert.ToString(values[1, column])?.Trim();
                if (string.IsNullOrWhiteSpace(header))
                    throw new ArgumentException("看板源区域的标题行不能包含空白字段名。");
                if (!headers.Add(header))
                    throw new ArgumentException("看板源区域包含重复字段名：「" + header + "」。");
            }

            foreach (var field in EnumerateRequestedFields())
                if (!headers.Contains(field))
                    throw new ArgumentException("源区域中找不到字段「" + field + "」。");

            var valueColumn = FindColumn(values, columnCount, _valueField);
            var hasNumericValue = false;
            for (var row = 2; row <= rowCount; row++)
            {
                double ignored;
                if (TryConvertDouble(values[row, valueColumn], out ignored))
                {
                    hasNumericValue = true;
                    break;
                }
            }
            if (!hasNumericValue)
                throw new ArgumentException("数值字段「" + _valueField + "」没有可用于看板聚合的数字。");
        }

        private IEnumerable<string> EnumerateRequestedFields()
        {
            yield return _categoryField;
            yield return _valueField;
            if (!string.IsNullOrWhiteSpace(_dateField)) yield return _dateField;
            if (!string.IsNullOrWhiteSpace(_seriesField)) yield return _seriesField;
            foreach (var field in _filterFields) yield return field;
        }

        private void BuildDashboardCanvas(Worksheet sheet, Worksheet sourceSheet, Range sourceRange)
        {
            var canvas = sheet.Range["A1:N100"];
            canvas.Interior.Color = ToOle("#F5F8F6");
            canvas.Font.Name = "微软雅黑";

            for (var column = 1; column <= 14; column++)
                ((Range)sheet.Columns[column]).ColumnWidth = 11;
            ((Range)sheet.Columns[1]).ColumnWidth = 14;
            ((Range)sheet.Columns[14]).ColumnWidth = 14;

            var titleRange = sheet.Range["A1:N2"];
            titleRange.Merge();
            titleRange.Value2 = _title;
            titleRange.Font.Name = "微软雅黑";
            titleRange.Font.Size = 24;
            titleRange.Font.Bold = true;
            titleRange.Font.Color = ToOle("#16372E");
            titleRange.VerticalAlignment = XlVAlign.xlVAlignCenter;

            var noteRange = sheet.Range["A3:N3"];
            noteRange.Merge();
            noteRange.Value2 = $"实时透视看板 · 来源：{sourceSheet.Name}!{sourceRange.Address} · 数据更新后使用“全部刷新”同步";
            noteRange.Font.Size = 10;
            noteRange.Font.Color = ToOle("#687A73");

            sheet.Range["A5:N5"].Merge();
            sheet.Range["A5"].Value2 = _filterFields.Length == 0 ? "概览" : "全局筛选";
            sheet.Range["A5"].Font.Size = 11;
            sheet.Range["A5"].Font.Bold = true;
            sheet.Range["A5"].Font.Color = ToOle("#476159");

            CreateCardFrame(sheet, "A10:D14", "A10:D10", "合计 " + _valueField, "#DDF2E8");
            CreateCardFrame(sheet, "E10:H14", "E10:H10", "平均 " + _valueField, "#E3EFFB");
            CreateCardFrame(sheet, "I10:L14", "I10:L10", "有效记录数", "#FFF0D8");

            var detailTitle = sheet.Range["A57:N58"];
            detailTitle.Merge();
            detailTitle.Value2 = "明细透视 · 可继续下钻和筛选";
            detailTitle.Font.Size = 14;
            detailTitle.Font.Bold = true;
            detailTitle.Font.Color = ToOle("#16372E");
        }

        private static void CreateCardFrame(Worksheet sheet, string cardAddress, string labelAddress, string label, string accentColor)
        {
            var card = sheet.Range[cardAddress];
            card.Interior.Color = ToOle("#FFFFFF");
            card.Borders.Color = ToOle("#DCE6E1");
            card.Borders.Weight = XlBorderWeight.xlThin;

            var labelRange = sheet.Range[labelAddress];
            labelRange.Merge();
            labelRange.Value2 = label;
            labelRange.Font.Size = 10;
            labelRange.Font.Bold = true;
            labelRange.Font.Color = ToOle("#4C635B");
            labelRange.Interior.Color = ToOle(accentColor);
            labelRange.HorizontalAlignment = XlHAlign.xlHAlignLeft;
            labelRange.VerticalAlignment = XlVAlign.xlVAlignCenter;
        }

        private PivotTable BuildMetricPivot(PivotCache cache, Worksheet sheet, string address, string name,
            string caption, XlConsolidationFunction function, string numberFormat)
        {
            var pivot = cache.CreatePivotTable(sheet.Range[address], UniqueName(name), Type.Missing, Type.Missing);
            var sourceField = (PivotField)pivot.PivotFields(_valueField);
            var dataField = pivot.AddDataField(sourceField, caption + " " + _valueField, function);
            dataField.NumberFormat = numberFormat;
            ConfigurePivot(pivot);
            return pivot;
        }

        private PivotTable BuildBreakdownPivot(PivotCache cache, Worksheet sheet, string address, string name,
            string dimension, string caption, XlConsolidationFunction function, bool sortDescending, int topN)
        {
            var pivot = cache.CreatePivotTable(sheet.Range[address], UniqueName(name), Type.Missing, Type.Missing);
            var rowField = (PivotField)pivot.PivotFields(dimension);
            rowField.Orientation = XlPivotFieldOrientation.xlRowField;
            rowField.Position = 1;
            var valueField = pivot.AddDataField((PivotField)pivot.PivotFields(_valueField), caption, function);
            valueField.NumberFormat = _numberFormat;
            ConfigurePivot(pivot);
            if (sortDescending)
                Try(() => rowField.AutoSort((int)XlSortOrder.xlDescending, valueField.Name));
            if (topN > 0)
                Try(() => ((dynamic)rowField).AutoShow(1, 1, topN, valueField.Name));
            TryHideBlankItem(rowField);
            pivot.RefreshTable();
            return pivot;
        }

        private PivotTable BuildDetailPivot(PivotCache cache, Worksheet sheet, string address, string name)
        {
            var pivot = cache.CreatePivotTable(sheet.Range[address], UniqueName(name), Type.Missing, Type.Missing);
            var category = (PivotField)pivot.PivotFields(_categoryField);
            category.Orientation = XlPivotFieldOrientation.xlRowField;
            category.Position = 1;
            if (!string.IsNullOrWhiteSpace(_seriesField) &&
                !string.Equals(_seriesField, _categoryField, StringComparison.OrdinalIgnoreCase))
            {
                var series = (PivotField)pivot.PivotFields(_seriesField);
                series.Orientation = XlPivotFieldOrientation.xlRowField;
                series.Position = 2;
            }
            var value = pivot.AddDataField((PivotField)pivot.PivotFields(_valueField), "指标值", ParseAggregation(_aggregation));
            value.NumberFormat = _numberFormat;
            pivot.RowGrand = true;
            pivot.ColumnGrand = false;
            pivot.TableStyle2 = "PivotStyleMedium4";
            Try(() => pivot.RowAxisLayout(XlLayoutRowType.xlTabularRow));
            pivot.RefreshTable();
            return pivot;
        }

        private static void ConfigurePivot(PivotTable pivot)
        {
            pivot.RowGrand = false;
            pivot.ColumnGrand = false;
            pivot.TableStyle2 = "PivotStyleMedium4";
            Try(() => pivot.RowAxisLayout(XlLayoutRowType.xlTabularRow));
            pivot.RefreshTable();
        }

        private static void TryHideBlankItem(PivotField field)
        {
            Try(() => ((PivotItem)field.PivotItems("(blank)")).Visible = false);
            Try(() => ((PivotItem)field.PivotItems("(空白)")).Visible = false);
        }

        private static void BindMetricCard(Worksheet dashboardSheet, string address, PivotTable pivot, string numberFormat)
        {
            var valueRange = dashboardSheet.Range[address];
            valueRange.Merge();
            var pivotSheet = (Worksheet)pivot.Parent;
            var pivotRange = (Range)pivot.TableRange1;
            var totalCell = (Range)pivotRange.Cells[pivotRange.Rows.Count, pivotRange.Columns.Count];
            valueRange.Formula = "='" + pivotSheet.Name.Replace("'", "''") + "'!" + totalCell.Address;
            valueRange.NumberFormat = numberFormat;
            valueRange.Font.Name = "微软雅黑";
            valueRange.Font.Size = 25;
            valueRange.Font.Bold = true;
            valueRange.Font.Color = ToOle("#16372E");
            valueRange.Interior.Color = ToOle("#FFFFFF");
            valueRange.HorizontalAlignment = XlHAlign.xlHAlignLeft;
            valueRange.VerticalAlignment = XlVAlign.xlVAlignCenter;
        }

        private static void CreateChart(Worksheet sheet, PivotTable pivot, string topLeftAddress, string bottomRightAddress,
            XlChartType chartType, string title, ChartStyle style)
        {
            sheet.Activate();
            sheet.Range["N1"].Select();
            var topLeft = sheet.Range[topLeftAddress];
            var bottomRight = sheet.Range[bottomRightAddress];
            var width = Convert.ToDouble(bottomRight.Left) + Convert.ToDouble(bottomRight.Width) - Convert.ToDouble(topLeft.Left);
            var height = Convert.ToDouble(bottomRight.Top) + Convert.ToDouble(bottomRight.Height) - Convert.ToDouble(topLeft.Top);
            var chartObject = ((ChartObjects)sheet.ChartObjects(Type.Missing)).Add(
                Convert.ToDouble(topLeft.Left), Convert.ToDouble(topLeft.Top), width, height);
            var chart = chartObject.Chart;
            chart.SetSourceData(pivot.TableRange1, XlRowCol.xlColumns);
            chart.ChartType = chartType;
            chart.PlotVisibleOnly = false;
            chart.HasTitle = true;
            chart.ChartTitle.Text = title;
            Try(() => ((dynamic)chart).ShowAllFieldButtons = false);

            Try(() =>
            {
                dynamic chartArea = chart.ChartArea;
                chartArea.Format.Fill.Solid();
                chartArea.Format.Fill.ForeColor.RGB = ToOle("#FFFFFF");
                chartArea.Format.Line.ForeColor.RGB = ToOle("#DCE6E1");
                chartArea.Format.Line.Weight = 0.75f;
            });
            Try(() =>
            {
                dynamic plotArea = chart.PlotArea;
                plotArea.Format.Fill.Solid();
                plotArea.Format.Fill.ForeColor.RGB = ToOle("#FFFFFF");
                plotArea.Format.Line.Visible = 0;
            });
            Try(() =>
            {
                dynamic titleFont = chart.ChartTitle.Format.TextFrame2.TextRange.Font;
                titleFont.Name = "微软雅黑";
                titleFont.Size = 15;
                titleFont.Bold = -1;
                titleFont.Fill.ForeColor.RGB = ToOle("#16372E");
            });

            if (style == ChartStyle.Share)
                StyleShareChart(chart);
            else
                StyleAxisChart(chart, style);
        }

        private static void StyleAxisChart(Microsoft.Office.Interop.Excel.Chart chart, ChartStyle style)
        {
            chart.HasLegend = false;
            Try(() =>
            {
                dynamic series = chart.SeriesCollection(1);
                var color = style == ChartStyle.Trend ? "#2F80ED" : style == ChartStyle.Comparison ? "#5C8FD6" : "#168653";
                if (style == ChartStyle.Trend)
                {
                    series.Format.Line.ForeColor.RGB = ToOle(color);
                    series.Format.Line.Weight = 2.75f;
                    series.MarkerStyle = XlMarkerStyle.xlMarkerStyleCircle;
                    series.MarkerSize = 7;
                    series.MarkerBackgroundColor = ToOle("#FFFFFF");
                    series.MarkerForegroundColor = ToOle(color);
                }
                else
                {
                    series.Format.Fill.Solid();
                    series.Format.Fill.ForeColor.RGB = ToOle(color);
                    series.Format.Line.Visible = 0;
                    series.ApplyDataLabels();
                    series.DataLabels(Type.Missing).Position = XlDataLabelPosition.xlLabelPositionOutsideEnd;
                    series.DataLabels(Type.Missing).NumberFormat = "#,##0.##";
                    ((dynamic)chart.ChartGroups(1)).GapWidth = 48;
                }
            });
            Try(() =>
            {
                dynamic categoryAxis = chart.Axes(XlAxisType.xlCategory, XlAxisGroup.xlPrimary);
                categoryAxis.TickLabels.Font.Name = "微软雅黑";
                categoryAxis.TickLabels.Font.Size = 9;
                categoryAxis.TickLabels.Font.Color = ToOle("#53645F");
                categoryAxis.Format.Line.ForeColor.RGB = ToOle("#DCE6E1");
            });
            Try(() =>
            {
                dynamic valueAxis = chart.Axes(XlAxisType.xlValue, XlAxisGroup.xlPrimary);
                valueAxis.TickLabels.Font.Name = "微软雅黑";
                valueAxis.TickLabels.Font.Size = 9;
                valueAxis.TickLabels.Font.Color = ToOle("#53645F");
                valueAxis.Format.Line.Visible = 0;
                valueAxis.MajorGridlines.Format.Line.ForeColor.RGB = ToOle("#E9EFEC");
            });
        }

        private static void StyleShareChart(Microsoft.Office.Interop.Excel.Chart chart)
        {
            chart.HasLegend = false;
            Try(() => ((dynamic)chart).DoughnutHoleSize = 66);
            Try(() => ((dynamic)chart).FirstSliceAngle = 270);
            var palette = new[] { "#0F5B4B", "#1B7F67", "#36A486", "#F0B35A", "#E97862", "#5C8FD6", "#9A7BCB" };
            Try(() =>
            {
                dynamic series = chart.SeriesCollection(1);
                for (var i = 1; i <= series.Points().Count; i++)
                {
                    dynamic point = series.Points(i);
                    point.Format.Fill.Solid();
                    point.Format.Fill.ForeColor.RGB = ToOle(palette[(i - 1) % palette.Length]);
                    point.Format.Line.ForeColor.RGB = ToOle("#FFFFFF");
                    point.Format.Line.Weight = 1.25f;
                }
                series.ApplyDataLabels();
                dynamic labels = series.DataLabels(Type.Missing);
                labels.ShowCategoryName = true;
                labels.ShowPercentage = true;
                labels.ShowValue = false;
                labels.ShowLegendKey = false;
                labels.Separator = "\n";
                labels.Font.Name = "微软雅黑";
                labels.Font.Size = 9;
            });
        }

        private static void PrepareDropdownPivotFields(IList<PivotTable> pivots, IEnumerable<string> filterFields)
        {
            foreach (var pivot in pivots)
            {
                foreach (var fieldName in filterFields)
                {
                    try
                    {
                        var field = (PivotField)pivot.PivotFields(fieldName);
                        if (field.Orientation == XlPivotFieldOrientation.xlHidden)
                        {
                            field.Orientation = XlPivotFieldOrientation.xlPageField;
                            field.EnableMultiplePageItems = false;
                        }
                    }
                    catch (Exception ex)
                    {
                        throw new InvalidOperationException(
                            "无法为下拉兼容筛选准备字段「" + fieldName + "」（透视表「" + pivot.Name + "」）。", ex);
                    }
                }
                pivot.RefreshTable();
            }
        }

        private static string CreateDropdownFilters(
            Microsoft.Office.Interop.Excel.Application application,
            Workbook workbook,
            Worksheet dashboardSheet,
            Worksheet supportSheet,
            Range sourceRange,
            IList<PivotTable> pivots,
            IList<string> filterFields)
        {
            const int listStartColumn = 52; // AZ 起，避开透视辅助区域。
            var bindingToken = Math.Abs(DateTime.Now.Ticks % 100000000).ToString(CultureInfo.InvariantCulture);
            var bindings = new List<DashboardFilterBinding>();
            // 筛选字段的去重取值统一从这一次读入的数组中获取,
            // 避免每个字段对整列逐行发起 COM 调用。
            var sourceValues = sourceRange.Value2 as object[,];
            if (sourceValues == null)
                throw new ArgumentException("看板源区域数据无法读取。");

            dashboardSheet.Range["A5:N5"].Value2 = "全局筛选 · 下拉兼容模式";
            for (var index = 0; index < filterFields.Count; index++)
            {
                var field = filterFields[index];
                var column = 1 + index * 4;
                var label = dashboardSheet.Range[
                    dashboardSheet.Cells[6, column], dashboardSheet.Cells[6, column + 3]];
                label.Merge();
                label.Value2 = field;
                label.Font.Name = "微软雅黑";
                label.Font.Size = 9;
                label.Font.Bold = true;
                label.Font.Color = ToOle("#476159");

                var control = dashboardSheet.Range[
                    dashboardSheet.Cells[7, column], dashboardSheet.Cells[8, column + 3]];
                control.Merge();
                control.Value2 = DashboardInteractionManager.AllSelection;
                control.Font.Name = "微软雅黑";
                control.Font.Size = 11;
                control.Font.Color = ToOle("#16372E");
                control.Interior.Color = ToOle("#FFFFFF");
                control.Borders.Color = ToOle("#BFD4CB");
                control.Borders.Weight = XlBorderWeight.xlThin;
                control.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                control.VerticalAlignment = XlVAlign.xlVAlignCenter;

                var values = ReadDistinctFilterValues(sourceValues, field);
                var listColumn = listStartColumn + index;
                ((Range)supportSheet.Cells[1, listColumn]).Value2 = field;
                ((Range)supportSheet.Cells[2, listColumn]).Value2 = DashboardInteractionManager.AllSelection;
                for (var row = 0; row < values.Count; row++)
                    ((Range)supportSheet.Cells[row + 3, listColumn]).Value2 = values[row];

                var listRange = supportSheet.Range[
                    supportSheet.Cells[2, listColumn], supportSheet.Cells[values.Count + 2, listColumn]];
                var listName = "AgentFilterList_" + bindingToken + "_" + index;
                workbook.Names.Add(listName, "='" + supportSheet.Name.Replace("'", "''") + "'!" + listRange.Address, true);
                control.Validation.Delete();
                control.Validation.Add(XlDVType.xlValidateList, XlDVAlertStyle.xlValidAlertStop,
                    XlFormatConditionOperator.xlBetween, "=" + listName, Type.Missing);
                control.Validation.IgnoreBlank = false;
                control.Validation.InCellDropdown = true;
                control.Validation.ErrorTitle = "无效筛选项";
                control.Validation.ErrorMessage = "请从下拉列表中选择。";
                control.Validation.ShowError = true;

                bindings.Add(new DashboardFilterBinding
                {
                    DashboardSheet = dashboardSheet.Name,
                    SupportSheet = supportSheet.Name,
                    ControlAddress = control.Address,
                    FieldName = field,
                    ListName = listName
                });
            }

            var bindingName = DashboardInteractionManager.RegisterDashboard(
                application, workbook, supportSheet, bindingToken, bindings);
            return bindingName;
        }

        private static List<object> ReadDistinctFilterValues(object[,] sourceValues, string field)
        {
            var columnCount = sourceValues.GetLength(1);
            var rowCount = sourceValues.GetLength(0);
            var column = FindColumn(sourceValues, columnCount, field);
            var values = new List<object>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (var row = 2; row <= rowCount; row++)
            {
                var raw = sourceValues[row, column];
                if (raw == null) continue;
                var key = Convert.ToString(raw, CultureInfo.InvariantCulture)?.Trim();
                if (string.IsNullOrWhiteSpace(key) || !seen.Add(key)) continue;
                values.Add(raw);
            }
            values.Sort((left, right) => string.Compare(
                Convert.ToString(left, CultureInfo.CurrentCulture),
                Convert.ToString(right, CultureInfo.CurrentCulture),
                StringComparison.CurrentCultureIgnoreCase));
            return values;
        }

        private static dynamic CreateSlicer(Workbook workbook, Worksheet dashboardSheet,
            IList<PivotTable> pivots, string field, int index)
        {
            dynamic caches = workbook.SlicerCaches;
            var seedIndex = Math.Min(3, pivots.Count - 1);
            dynamic seedPivot = pivots[seedIndex];
            dynamic seedField = seedPivot.PivotFields(field);
            var suffix = Math.Abs(DateTime.Now.Ticks % 100000000).ToString(CultureInfo.InvariantCulture) + index;
            var cacheName = "AgentSlicerCache" + suffix;
            dynamic slicerCache;
            try { slicerCache = caches.Add(seedPivot, field, cacheName, Type.Missing); }
            catch { slicerCache = caches.Add2(seedPivot, seedField, cacheName, Type.Missing); }

            for (var i = 0; i < pivots.Count; i++)
            {
                if (i == seedIndex) continue;
                if (IsPivotConnected(slicerCache, pivots[i])) continue;
                if (Convert.ToInt32(pivots[i].CacheIndex) == Convert.ToInt32(seedPivot.CacheIndex))
                    continue;
                try
                {
                    pivots[i].ChangePivotCache((PivotCache)seedPivot.PivotCache());
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException(
                        "无法把全局筛选字段「" + field + "」连接到透视表「" + pivots[i].Name +
                        "」（主缓存 " + Convert.ToString(seedPivot.CacheIndex) + "，目标缓存 " + pivots[i].CacheIndex +
                        "，当前已连接 " + DescribeConnectedPivots(slicerCache) + "）。", ex);
                }
            }
            System.Windows.Forms.Application.DoEvents();

            var left = Convert.ToDouble(((Range)dashboardSheet.Cells[6, 1 + index * 4]).Left);
            var top = Convert.ToDouble(((Range)dashboardSheet.Cells[6, 1]).Top);
            var width = Convert.ToDouble(dashboardSheet.Range[dashboardSheet.Cells[6, 1 + index * 4], dashboardSheet.Cells[6, Math.Min(4 + index * 4, 14)]].Width);
            var height = Convert.ToDouble(dashboardSheet.Range[dashboardSheet.Cells[6, 1], dashboardSheet.Cells[8, 1]].Height);
            dynamic slicer = slicerCache.Slicers.Add(dashboardSheet, Type.Missing,
                "AgentSlicer" + suffix, field, top, left, width, height);
            Try(() => slicer.Style = "SlicerStyleLight2");
            Try(() => slicer.NumberOfColumns = 2);
            return slicerCache;
        }

        private static bool IsPivotConnected(dynamic slicerCache, PivotTable pivot)
        {
            dynamic connected = slicerCache.PivotTables;
            for (var i = 1; i <= Convert.ToInt32(connected.Count); i++)
            {
                dynamic item = connected.Item(i);
                if (string.Equals(Convert.ToString(item.Name), pivot.Name, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            return false;
        }

        private static void ConnectAllSlicerCaches(IList<dynamic> slicerCaches, IList<PivotTable> pivots)
        {
            System.Windows.Forms.Application.DoEvents();
            foreach (dynamic slicerCache in slicerCaches)
            {
                foreach (var pivot in pivots)
                {
                    if (IsPivotConnected(slicerCache, pivot)) continue;
                    try { slicerCache.PivotTables.AddPivotTable(pivot); }
                    catch
                    {
                        System.Windows.Forms.Application.DoEvents();
                        if (!IsPivotConnected(slicerCache, pivot)) throw;
                    }
                }
            }
        }

        private static string DescribeConnectedPivots(dynamic slicerCache)
        {
            var names = new List<string>();
            dynamic connected = slicerCache.PivotTables;
            for (var i = 1; i <= Convert.ToInt32(connected.Count); i++)
            {
                try { names.Add(Convert.ToString(connected.Item(i).Name)); }
                catch { names.Add("?"); }
            }
            return Convert.ToString(connected.Count) + " 张 [" + string.Join(",", names) + "]";
        }

        private static int FindColumn(object[,] values, int columnCount, string field)
        {
            for (var column = 1; column <= columnCount; column++)
            {
                var header = Convert.ToString(values[1, column])?.Trim();
                if (string.Equals(header, field, StringComparison.OrdinalIgnoreCase)) return column;
            }
            return -1;
        }

        private static bool TryConvertDouble(object value, out double number)
        {
            if (value == null)
            {
                number = 0;
                return false;
            }
            if (value is double)
            {
                number = (double)value;
                return true;
            }
            return double.TryParse(Convert.ToString(value), NumberStyles.Any, CultureInfo.CurrentCulture, out number) ||
                   double.TryParse(Convert.ToString(value), NumberStyles.Any, CultureInfo.InvariantCulture, out number);
        }

        private static XlConsolidationFunction ParseAggregation(string value)
        {
            switch ((value ?? "sum").Trim().ToLowerInvariant())
            {
                case "sum": return XlConsolidationFunction.xlSum;
                case "count": return XlConsolidationFunction.xlCount;
                case "average": return XlConsolidationFunction.xlAverage;
                default: throw new ArgumentException("aggregation 仅支持 sum、count 或 average。");
            }
        }

        private static string UniqueName(string prefix)
        {
            return prefix + Math.Abs(DateTime.Now.Ticks % 100000000).ToString(CultureInfo.InvariantCulture);
        }

        private static int ToOle(string htmlColor)
        {
            return ColorTranslator.ToOle(ColorTranslator.FromHtml(htmlColor));
        }

        private static void Try(System.Action action)
        {
            try { action(); } catch { }
        }

        private enum ChartStyle
        {
            Ranking,
            Trend,
            Share,
            Comparison
        }

        public sealed class Factory : IOperationFactory
        {
            public string ToolName => "dashboard_create";

            public IOperation Parse(string argumentsJson)
            {
                using (var document = JsonDocument.Parse(argumentsJson))
                {
                    var root = document.RootElement;
                    var sourceAddress = ReadString(root, "source_address");
                    var categoryField = ReadString(root, "category_field");
                    var valueField = ReadString(root, "value_field");
                    var aggregation = (ReadString(root, "aggregation") ?? "sum").Trim().ToLowerInvariant();
                    var filterMode = (ReadString(root, "filter_mode") ?? "auto").Trim().ToLowerInvariant();
                    var topN = ReadInt(root, "top_n", 10);
                    var filterFields = ReadStringArray(root, "filter_fields");

                    if (string.IsNullOrWhiteSpace(sourceAddress)) throw new ArgumentException("source_address 不能为空。");
                    if (string.IsNullOrWhiteSpace(categoryField)) throw new ArgumentException("category_field 不能为空。");
                    if (string.IsNullOrWhiteSpace(valueField)) throw new ArgumentException("value_field 不能为空。");
                    ParseAggregation(aggregation);
                    if (filterMode != "auto" && filterMode != "slicer" && filterMode != "dropdown")
                        throw new ArgumentException("filter_mode 仅支持 auto、slicer 或 dropdown。");
                    if (topN < 3 || topN > 20) throw new ArgumentException("top_n 必须在 3 到 20 之间。");
                    if (filterFields.Length > 3) throw new ArgumentException("filter_fields 最多支持 3 个全局筛选字段。");

                    return new CreateDashboardOp(
                        ReadString(root, "source_sheet"), sourceAddress,
                        ReadString(root, "dashboard_sheet_name") ?? "Agent看板",
                        ReadString(root, "title") ?? "业务分析看板",
                        ReadString(root, "date_field"), categoryField,
                        ReadString(root, "series_field"), valueField,
                        filterFields, filterMode, aggregation, topN,
                        ReadString(root, "number_format") ?? "#,##0.##");
                }
            }

            private static string ReadString(JsonElement root, string name)
            {
                return root.TryGetProperty(name, out var value) && value.ValueKind == JsonValueKind.String
                    ? value.GetString()?.Trim()
                    : null;
            }

            private static int ReadInt(JsonElement root, string name, int fallback)
            {
                if (!root.TryGetProperty(name, out var value)) return fallback;
                if (value.ValueKind != JsonValueKind.Number || !value.TryGetInt32(out var result))
                    throw new ArgumentException(name + " 必须是整数。");
                return result;
            }

            private static string[] ReadStringArray(JsonElement root, string name)
            {
                if (!root.TryGetProperty(name, out var value)) return new string[0];
                if (value.ValueKind != JsonValueKind.Array)
                    throw new ArgumentException(name + " 必须是字符串数组。");
                var result = new List<string>();
                var unique = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var item in value.EnumerateArray())
                {
                    if (item.ValueKind != JsonValueKind.String || string.IsNullOrWhiteSpace(item.GetString()))
                        throw new ArgumentException(name + " 不能包含空值。");
                    var text = item.GetString().Trim();
                    if (unique.Add(text)) result.Add(text);
                }
                return result.ToArray();
            }
        }
    }
}
