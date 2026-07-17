using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Text.Json;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;

namespace AgentForExcel.Operations.Chart
{
    /// <summary>根据工作表区域创建现代化嵌入式图表；默认在新的分析页中生成值快照。</summary>
    public sealed class CreateChartOp : IOperation
    {
        public string ToolName => "chart_create";
        public bool IsWriteOperation => true;

        private readonly string _sourceSheet;
        private readonly string _sourceAddress;
        private readonly string _destinationSheet;
        private readonly string _anchorAddress;
        private readonly string _chartType;
        private readonly string _title;
        private readonly string _name;
        private readonly double _width;
        private readonly double _height;
        private readonly string _categoryField;
        private readonly string _valueField;
        private readonly HashSet<string> _excludeCategories;
        private readonly bool _sortDescending;
        private readonly bool _showDataLabels;
        private readonly bool _showPercentage;
        private readonly string _legendPosition;
        private readonly string _palette;
        private readonly string _aggregation;
        private readonly int _maxCategories;
        private readonly bool _includeOther;

        private CreateChartOp(
            string sourceSheet, string sourceAddress, string destinationSheet, string anchorAddress,
            string chartType, string title, string name, double width, double height,
            string categoryField, string valueField, HashSet<string> excludeCategories,
            bool sortDescending, bool showDataLabels, bool showPercentage,
            string legendPosition, string palette, string aggregation, int maxCategories, bool includeOther)
        {
            _sourceSheet = sourceSheet;
            _sourceAddress = sourceAddress;
            _destinationSheet = destinationSheet;
            _anchorAddress = anchorAddress;
            _chartType = chartType;
            _title = title;
            _name = name;
            _width = width;
            _height = height;
            _categoryField = categoryField;
            _valueField = valueField;
            _excludeCategories = excludeCategories ?? new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            _sortDescending = sortDescending;
            _showDataLabels = showDataLabels;
            _showPercentage = showPercentage;
            _legendPosition = legendPosition;
            _palette = palette;
            _aggregation = aggregation;
            _maxCategories = maxCategories;
            _includeOther = includeOther;
        }

        public string Describe()
        {
            var source = (string.IsNullOrWhiteSpace(_sourceSheet) ? "活动工作表" : "工作表「" + _sourceSheet + "」") + "!" + _sourceAddress;
            var destination = string.IsNullOrWhiteSpace(_destinationSheet)
                ? "新的安全图表工作表"
                : "工作表「" + _destinationSheet + "」";
            return $"根据 {source} 创建现代 {_chartType} 图表，放置在{destination}；原始数据不会被修改";
        }

        public string Execute(AppContext context)
        {
            var sourceSheet = Cell.CellOperationSupport.GetWorksheet(context, _sourceSheet);
            var sourceRange = Cell.CellOperationSupport.GetRange(sourceSheet, _sourceAddress);
            if (sourceRange.Rows.Count < 2 || sourceRange.Columns.Count < 2)
                throw new ArgumentException("图表源区域至少需要 2 行 × 2 列，并包含标题行。");

            var isProportion = IsProportionChart(_chartType);
            var useExplicitSeries = isProportion ||
                                    !string.IsNullOrWhiteSpace(_categoryField) ||
                                    !string.IsNullOrWhiteSpace(_valueField) ||
                                    _excludeCategories.Count > 0;
            var seriesData = useExplicitSeries ? ReadSeriesData(sourceRange, isProportion) : null;

            Worksheet destinationSheet = null;
            Worksheet createdSheet = null;
            ChartObject chartObject = null;
            try
            {
                Range chartSourceRange = sourceRange;
                if (string.IsNullOrWhiteSpace(_destinationSheet))
                {
                    createdSheet = Analysis.AnalysisSheetSupport.CreateUniqueWorksheet(context, "Agent图表");
                    destinationSheet = createdSheet;
                    chartSourceRange = BuildAnalysisSheet(destinationSheet, sourceSheet, sourceRange, seriesData);
                }
                else
                {
                    destinationSheet = Cell.CellOperationSupport.GetWorksheet(context, _destinationSheet);
                }

                var anchor = ResolveAnchor(destinationSheet, chartSourceRange, createdSheet != null);
                var chartObjects = (ChartObjects)destinationSheet.ChartObjects(Type.Missing);
                chartObject = chartObjects.Add(
                    Convert.ToDouble(anchor.Left), Convert.ToDouble(anchor.Top), _width, _height);
                chartObject.Name = MakeUniqueName(chartObjects, _name);

                var chart = chartObject.Chart;
                chart.ChartType = ParseChartType(_chartType);
                Series explicitSeries = null;
                if (seriesData != null)
                    explicitSeries = createdSheet != null
                        ? BindSnapshotSeries(chart, chartSourceRange)
                        : BindExplicitSeries(chart, seriesData);
                else
                    chart.SetSourceData(chartSourceRange, XlRowCol.xlColumns);

                chart.HasTitle = !string.IsNullOrWhiteSpace(_title);
                if (chart.HasTitle) chart.ChartTitle.Text = _title;
                ApplyModernStyle(chart, explicitSeries, seriesData, isProportion);

                var processingNote = seriesData == null || string.IsNullOrWhiteSpace(seriesData.ProcessingNote)
                    ? ""
                    : "；" + seriesData.ProcessingNote;
                var sourceNote = createdSheet == null
                    ? ""
                    : "，并在该页保留了用于展示的值快照；原始工作表未修改";
                return $"已创建现代图表「{chartObject.Name}」，位置为 {destinationSheet.Name}!{anchor.Address}{sourceNote}{processingNote}。";
            }
            catch
            {
                try { chartObject?.Delete(); } catch { }
                if (createdSheet != null)
                    Analysis.AnalysisSheetSupport.DeleteWorksheetSilently(context, createdSheet);
                throw;
            }
        }

        private Range BuildAnalysisSheet(Worksheet destinationSheet, Worksheet sourceSheet, Range sourceRange, SeriesData seriesData)
        {
            var titleCell = (Range)destinationSheet.Cells[1, 1];
            titleCell.Value2 = string.IsNullOrWhiteSpace(_title) ? "Agent 分析图表" : _title;
            titleCell.Font.Name = "微软雅黑";
            titleCell.Font.Size = 20;
            titleCell.Font.Bold = true;
            titleCell.Font.Color = ToOle("#18352D");

            var noteCell = (Range)destinationSheet.Cells[2, 1];
            noteCell.Value2 = $"数据快照 · 来源：{sourceSheet.Name}!{sourceRange.Address} · 原表未修改";
            noteCell.Font.Name = "微软雅黑";
            noteCell.Font.Size = 10;
            noteCell.Font.Color = ToOle("#6B7C76");

            Range snapshot;
            if (seriesData != null)
            {
                var values = new object[seriesData.Categories.Count + 1, 2];
                values[0, 0] = seriesData.CategoryHeader;
                values[0, 1] = seriesData.ValueHeader;
                for (var row = 0; row < seriesData.Categories.Count; row++)
                {
                    values[row + 1, 0] = seriesData.Categories[row];
                    values[row + 1, 1] = seriesData.Values[row];
                }
                snapshot = destinationSheet.Range[
                    destinationSheet.Cells[4, 1],
                    destinationSheet.Cells[seriesData.Categories.Count + 4, 2]];
                snapshot.Value2 = values;
            }
            else
            {
                snapshot = destinationSheet.Range[
                    destinationSheet.Cells[4, 1],
                    destinationSheet.Cells[sourceRange.Rows.Count + 3, sourceRange.Columns.Count]];
                snapshot.Value2 = sourceRange.Value2;
                try { snapshot.NumberFormat = sourceRange.NumberFormat; } catch { }
            }

            var table = destinationSheet.ListObjects.Add(
                XlListObjectSourceType.xlSrcRange,
                snapshot,
                Type.Missing,
                XlYesNoGuess.xlYes,
                Type.Missing);
            table.Name = "AgentChartData" + (DateTime.Now.Ticks % 100000000).ToString(CultureInfo.InvariantCulture);
            table.TableStyle = "TableStyleMedium4";
            snapshot.Columns.AutoFit();
            for (var column = 1; column <= snapshot.Columns.Count; column++)
            {
                var targetColumn = (Range)snapshot.Columns[column];
                if (Convert.ToDouble(targetColumn.ColumnWidth) > 24) targetColumn.ColumnWidth = 24;
            }
            destinationSheet.Tab.Color = ToOle("#168653");
            return snapshot;
        }

        private Range ResolveAnchor(Worksheet destinationSheet, Range chartSourceRange, bool isNewSheet)
        {
            if (!string.IsNullOrWhiteSpace(_anchorAddress))
                return Cell.CellOperationSupport.GetRange(destinationSheet, _anchorAddress);
            if (!isNewSheet)
                return Cell.CellOperationSupport.GetRange(destinationSheet, "F2");

            var column = Math.Max(4, Math.Min(chartSourceRange.Columns.Count + 2, 12));
            return (Range)destinationSheet.Cells[4, column];
        }

        private SeriesData ReadSeriesData(Range sourceRange, bool isProportion)
        {
            var categoryColumn = ResolveColumn(sourceRange, _categoryField, false);
            var valueColumn = ResolveColumn(sourceRange, _valueField, true);
            if (categoryColumn == valueColumn)
                throw new ArgumentException("分类字段和数值字段不能是同一列。");

            var result = new SeriesData
            {
                CategoryHeader = Convert.ToString(((Range)sourceRange.Cells[1, categoryColumn]).Value2)?.Trim() ?? "分类",
                ValueHeader = Convert.ToString(((Range)sourceRange.Cells[1, valueColumn]).Value2)?.Trim() ?? "数值",
                OriginalCount = sourceRange.Rows.Count - 1
            };

            for (var row = 2; row <= sourceRange.Rows.Count; row++)
            {
                var category = Convert.ToString(((Range)sourceRange.Cells[row, categoryColumn]).Value2)?.Trim();
                if (string.IsNullOrWhiteSpace(category) || ShouldExclude(category)) continue;
                object rawValue = ((Range)sourceRange.Cells[row, valueColumn]).Value2;
                double value;
                if (!TryConvertDouble(rawValue, out value)) continue;
                if (isProportion && value < 0)
                    throw new ArgumentException("占比图不能包含负数，请先明确负值的处理方式。");
                if (isProportion && Math.Abs(value) < 0.0000001) continue;
                result.Categories.Add(category);
                result.Values.Add(value);
            }

            if (result.Categories.Count == 0)
                throw new ArgumentException("没有找到可用于图表的有效分类和数值。");
            ApplyAggregation(result);
            ApplyCategoryReduction(result, isProportion);
            if (IsOrderedChart(_chartType) && !_sortDescending)
                SortSeriesChronologically(result);
            if (_sortDescending)
                SortSeriesDescending(result);
            return result;
        }

        private void ApplyAggregation(SeriesData data)
        {
            var effective = (_aggregation ?? "auto").Trim().ToLowerInvariant();
            var duplicateCount = data.Categories.Count - new HashSet<string>(data.Categories, StringComparer.OrdinalIgnoreCase).Count;
            if (effective == "auto") effective = duplicateCount > 0 ? "sum" : "none";
            if (effective == "none") return;

            var order = new List<string>();
            var groups = new Dictionary<string, AggregateBucket>(StringComparer.OrdinalIgnoreCase);
            for (var i = 0; i < data.Categories.Count; i++)
            {
                AggregateBucket bucket;
                if (!groups.TryGetValue(data.Categories[i], out bucket))
                {
                    bucket = new AggregateBucket { Label = data.Categories[i], Min = data.Values[i], Max = data.Values[i] };
                    groups[data.Categories[i]] = bucket;
                    order.Add(data.Categories[i]);
                }
                bucket.Sum += data.Values[i];
                bucket.Count++;
                if (data.Values[i] < bucket.Min) bucket.Min = data.Values[i];
                if (data.Values[i] > bucket.Max) bucket.Max = data.Values[i];
            }

            data.Categories.Clear();
            data.Values.Clear();
            foreach (var key in order)
            {
                var bucket = groups[key];
                data.Categories.Add(bucket.Label);
                switch (effective)
                {
                    case "average": data.Values.Add(bucket.Sum / bucket.Count); break;
                    case "count": data.Values.Add(bucket.Count); break;
                    case "min": data.Values.Add(bucket.Min); break;
                    case "max": data.Values.Add(bucket.Max); break;
                    default: data.Values.Add(bucket.Sum); break;
                }
            }
            data.ValueHeader = AggregationCaption(effective) + data.ValueHeader;
            data.ProcessingNote = $"已按「{data.CategoryHeader}」{AggregationCaption(effective)}，将 {data.OriginalCount} 条明细整理为 {data.Categories.Count} 个绘图点";
        }

        private void ApplyCategoryReduction(SeriesData data, bool isProportion)
        {
            var limit = isProportion ? Math.Min(8, _maxCategories) : _maxCategories;
            if (IsOrderedChart(_chartType) || data.Categories.Count <= limit) return;

            SortSeriesDescending(data);
            var keep = Math.Max(1, isProportion && _includeOther ? limit - 1 : limit);
            var other = 0d;
            for (var i = keep; i < data.Values.Count; i++) other += data.Values[i];
            if (data.Categories.Count > keep)
            {
                data.Categories.RemoveRange(keep, data.Categories.Count - keep);
                data.Values.RemoveRange(keep, data.Values.Count - keep);
                if (_includeOther && Math.Abs(other) > 0.0000001)
                {
                    data.Categories.Add("其他");
                    data.Values.Add(other);
                }
                var suffix = _includeOther ? "，长尾合并为“其他”" : "";
                data.ProcessingNote = AppendNote(data.ProcessingNote, $"分类数量已压缩到 {data.Categories.Count} 个{suffix}");
            }
        }

        private bool ShouldExclude(string category)
        {
            if (_excludeCategories.Contains(category)) return true;
            var normalized = category.Replace(" ", "").Trim();
            return string.Equals(normalized, "Total", StringComparison.OrdinalIgnoreCase) ||
                   normalized == "合计" || normalized == "总计" || normalized == "小计";
        }

        private static int ResolveColumn(Range sourceRange, string field, bool preferNumeric)
        {
            if (!string.IsNullOrWhiteSpace(field))
            {
                for (var column = 1; column <= sourceRange.Columns.Count; column++)
                {
                    var header = Convert.ToString(((Range)sourceRange.Cells[1, column]).Value2)?.Trim();
                    if (string.Equals(header, field.Trim(), StringComparison.OrdinalIgnoreCase)) return column;
                }
                throw new ArgumentException("找不到图表字段：「" + field + "」。");
            }

            if (!preferNumeric) return 1;
            for (var column = sourceRange.Columns.Count; column >= 1; column--)
                for (var row = 2; row <= sourceRange.Rows.Count; row++)
                {
                    double numericValue;
                    object rawValue = ((Range)sourceRange.Cells[row, column]).Value2;
                    if (TryConvertDouble(rawValue, out numericValue)) return column;
                }
            throw new ArgumentException("图表区域中没有可用的数值列。");
        }

        private static bool TryConvertDouble(object value, out double result)
        {
            if (value == null)
            {
                result = 0;
                return false;
            }
            try
            {
                result = Convert.ToDouble(value, CultureInfo.InvariantCulture);
                return !double.IsNaN(result) && !double.IsInfinity(result);
            }
            catch
            {
                result = 0;
                return false;
            }
        }

        private static void SortSeriesDescending(SeriesData data)
        {
            for (var i = 0; i < data.Values.Count - 1; i++)
            {
                for (var j = i + 1; j < data.Values.Count; j++)
                {
                    if (data.Values[j] <= data.Values[i]) continue;
                    var value = data.Values[i];
                    data.Values[i] = data.Values[j];
                    data.Values[j] = value;
                    var category = data.Categories[i];
                    data.Categories[i] = data.Categories[j];
                    data.Categories[j] = category;
                }
            }
        }

        private static void SortSeriesChronologically(SeriesData data)
        {
            var keyed = new List<ChronologicalPoint>();
            for (var i = 0; i < data.Categories.Count; i++)
            {
                long key;
                if (!TryTimeKey(data.Categories[i], out key)) return;
                keyed.Add(new ChronologicalPoint { Category = data.Categories[i], Value = data.Values[i], Key = key, Order = i });
            }
            keyed.Sort((left, right) =>
            {
                var comparison = left.Key.CompareTo(right.Key);
                return comparison != 0 ? comparison : left.Order.CompareTo(right.Order);
            });
            data.Categories.Clear();
            data.Values.Clear();
            foreach (var point in keyed)
            {
                data.Categories.Add(point.Category);
                data.Values.Add(point.Value);
            }
        }

        private static bool TryTimeKey(string category, out long key)
        {
            var text = (category ?? string.Empty).Trim();
            var quarter = Regex.Match(text, @"^(\d{4})[-/]?Q([1-4])$", RegexOptions.IgnoreCase);
            if (!quarter.Success)
                quarter = Regex.Match(text, @"^(\d{4})年(?:第)?([1234一二三四])季度$");
            if (quarter.Success)
            {
                var qText = quarter.Groups[2].Value;
                var q = qText == "一" ? 1 : qText == "二" ? 2 : qText == "三" ? 3 : qText == "四" ? 4 : Convert.ToInt32(qText);
                key = Convert.ToInt64(quarter.Groups[1].Value) * 100 + q * 3;
                return true;
            }

            DateTime date;
            if (DateTime.TryParse(text, CultureInfo.CurrentCulture, DateTimeStyles.None, out date) ||
                DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.None, out date))
            {
                key = date.Ticks;
                return true;
            }
            var yearMonth = Regex.Match(text, @"^(\d{4})[-/](\d{1,2})$");
            if (yearMonth.Success)
            {
                key = Convert.ToInt64(yearMonth.Groups[1].Value) * 100 + Convert.ToInt64(yearMonth.Groups[2].Value);
                return true;
            }
            key = 0;
            return false;
        }

        private static bool IsOrderedChart(string chartType)
        {
            var value = (chartType ?? string.Empty).Trim().ToLowerInvariant();
            return value == "line" || value == "area";
        }

        private static bool ShouldShowLabels(SeriesData data, string chartType)
        {
            if (data == null) return true;
            var value = (chartType ?? string.Empty).Trim().ToLowerInvariant();
            if (value == "line" || value == "area") return data.Categories.Count <= 12;
            if (value == "scatter") return data.Categories.Count <= 20;
            return data.Categories.Count <= 20;
        }

        private static XlDisplayUnit? DetermineDisplayUnit(IList<double> values)
        {
            var maximum = 0d;
            foreach (var value in values) maximum = Math.Max(maximum, Math.Abs(value));
            if (maximum >= 1000000000d) return XlDisplayUnit.xlThousandMillions;
            if (maximum >= 1000000d) return XlDisplayUnit.xlMillions;
            if (maximum >= 1000d) return XlDisplayUnit.xlThousands;
            return null;
        }

        private static string AggregationCaption(string aggregation)
        {
            switch ((aggregation ?? "sum").Trim().ToLowerInvariant())
            {
                case "average": return "平均";
                case "count": return "计数";
                case "min": return "最小";
                case "max": return "最大";
                default: return "汇总";
            }
        }

        private static string AppendNote(string existing, string note)
        {
            return string.IsNullOrWhiteSpace(existing) ? note : existing + "；" + note;
        }

        private static Series BindExplicitSeries(Microsoft.Office.Interop.Excel.Chart chart, SeriesData data)
        {
            var collection = (SeriesCollection)chart.SeriesCollection(Type.Missing);
            while (collection.Count > 0) ((Series)collection.Item(1)).Delete();
            var series = collection.NewSeries();
            series.Name = data.ValueHeader;
            var categories = new object[data.Categories.Count];
            var values = new object[data.Values.Count];
            for (var index = 0; index < data.Categories.Count; index++)
            {
                categories[index] = data.Categories[index];
                values[index] = data.Values[index];
            }
            series.XValues = categories;
            series.Values = values;
            return series;
        }

        private static Series BindSnapshotSeries(Microsoft.Office.Interop.Excel.Chart chart, Range snapshot)
        {
            chart.SetSourceData(snapshot, XlRowCol.xlColumns);
            var collection = (SeriesCollection)chart.SeriesCollection(Type.Missing);
            if (collection.Count < 1)
                throw new InvalidOperationException("无法从分析快照创建图表序列。");
            return (Series)collection.Item(1);
        }

        private void ApplyModernStyle(Microsoft.Office.Interop.Excel.Chart chart, Series explicitSeries, SeriesData data, bool isProportion)
        {
            var palette = GetPalette(_palette);
            Try(() =>
            {
                dynamic area = chart.ChartArea;
                area.Format.Fill.Solid();
                area.Format.Fill.ForeColor.RGB = ToOle("#FFFFFF");
                area.Format.Line.Visible = 0;
            });
            Try(() =>
            {
                dynamic area = chart.PlotArea;
                area.Format.Fill.Solid();
                area.Format.Fill.ForeColor.RGB = ToOle("#FFFFFF");
                area.Format.Line.Visible = 0;
            });
            Try(() =>
            {
                dynamic font = chart.ChartTitle.Format.TextFrame2.TextRange.Font;
                font.Name = "微软雅黑";
                font.Size = 18;
                font.Bold = -1;
                font.Fill.ForeColor.RGB = ToOle("#18352D");
            });

            if (isProportion && explicitSeries != null)
            {
                Try(() => ((dynamic)chart).DoughnutHoleSize = _chartType.Equals("doughnut", StringComparison.OrdinalIgnoreCase) ? 64 : 0);
                Try(() => ((dynamic)chart).FirstSliceAngle = 270);
                for (var index = 1; index <= data.Categories.Count; index++)
                {
                    var color = palette[(index - 1) % palette.Length];
                    Try(() =>
                    {
                        dynamic point = explicitSeries.Points(index);
                        point.Format.Fill.Solid();
                        point.Format.Fill.ForeColor.RGB = ToOle(color);
                        point.Format.Line.ForeColor.RGB = ToOle("#FFFFFF");
                        point.Format.Line.Weight = 1.5f;
                    });
                }

                if (_showDataLabels && data.Categories.Count <= 8)
                {
                    Try(() => explicitSeries.ApplyDataLabels());
                    Try(() =>
                    {
                        dynamic labels = explicitSeries.DataLabels(Type.Missing);
                        labels.ShowLegendKey = false;
                        labels.ShowCategoryName = true;
                        labels.ShowValue = !_showPercentage;
                        labels.ShowPercentage = _showPercentage;
                        labels.ShowSeriesName = false;
                        labels.Separator = "\n";
                        labels.Font.Name = "微软雅黑";
                        labels.Font.Size = 10;
                    });
                    var labelsOutside = TrySetOutsideLabels(explicitSeries, data.Categories.Count);
                    for (var index = 1; index <= data.Categories.Count; index++)
                    {
                        var color = labelsOutside ? "#29453D" : GetContrastTextColor(palette[(index - 1) % palette.Length]);
                        var labelIndex = index;
                        Try(() => ((dynamic)explicitSeries.DataLabels(labelIndex)).Font.Color = ToOle(color));
                    }
                    if (labelsOutside) Try(() => ((dynamic)chart).HasLeaderLines = true);
                }

                ApplyLegend(chart, _showDataLabels && data.Categories.Count <= 8 ? "none" : _legendPosition);
                if (_chartType.Equals("doughnut", StringComparison.OrdinalIgnoreCase))
                    AddCenterSummary(chart, data);
                return;
            }

            StyleStandardSeries(chart, palette, data);
            ApplyLegend(chart, _legendPosition);
            StyleAxes(chart, data);
        }

        private void StyleStandardSeries(Microsoft.Office.Interop.Excel.Chart chart, string[] palette, SeriesData data)
        {
            Try(() =>
            {
                var seriesCollection = (SeriesCollection)chart.SeriesCollection(Type.Missing);
                for (var index = 1; index <= seriesCollection.Count; index++)
                {
                    dynamic series = seriesCollection.Item(index);
                    var color = palette[(index - 1) % palette.Length];
                    switch ((_chartType ?? string.Empty).ToLowerInvariant())
                    {
                        case "column":
                        case "bar":
                            series.Format.Fill.Solid();
                            series.Format.Fill.ForeColor.RGB = ToOle(color);
                            series.Format.Line.Visible = 0;
                            break;
                        case "line":
                            series.Format.Line.ForeColor.RGB = ToOle(color);
                            series.Format.Line.Weight = 2.5f;
                            series.MarkerStyle = data != null && data.Categories.Count > 40
                                ? XlMarkerStyle.xlMarkerStyleNone
                                : XlMarkerStyle.xlMarkerStyleCircle;
                            series.MarkerSize = data != null && data.Categories.Count > 24 ? 5 : 7;
                            series.MarkerForegroundColor = ToOle(color);
                            series.MarkerBackgroundColor = ToOle("#FFFFFF");
                            break;
                        case "area":
                            series.Format.Fill.Solid();
                            series.Format.Fill.ForeColor.RGB = ToOle(color);
                            series.Format.Fill.Transparency = Math.Min(0.58, 0.28 + (index - 1) * 0.10);
                            series.Format.Line.ForeColor.RGB = ToOle(color);
                            series.Format.Line.Weight = 2f;
                            break;
                        case "scatter":
                            series.Format.Line.Visible = 0;
                            series.MarkerStyle = XlMarkerStyle.xlMarkerStyleCircle;
                            series.MarkerSize = 8;
                            series.MarkerForegroundColor = ToOle("#FFFFFF");
                            series.MarkerBackgroundColor = ToOle(color);
                            break;
                        default:
                            series.Format.Fill.Solid();
                            series.Format.Fill.ForeColor.RGB = ToOle(color);
                            series.Format.Line.ForeColor.RGB = ToOle(color);
                            series.Format.Line.Weight = 1.75f;
                            break;
                    }

                    if (_showDataLabels && ShouldShowLabels(data, _chartType))
                        ApplyStandardDataLabels(series, _chartType);
                }

                if (_chartType.Equals("column", StringComparison.OrdinalIgnoreCase) ||
                    _chartType.Equals("bar", StringComparison.OrdinalIgnoreCase))
                {
                    dynamic group = chart.ChartGroups(1);
                    group.GapWidth = 62;
                    group.Overlap = 0;
                }
            });
        }

        private static void ApplyStandardDataLabels(dynamic series, string chartType)
        {
            Try(() => series.ApplyDataLabels());
            Try(() =>
            {
                dynamic labels = series.DataLabels(Type.Missing);
                labels.ShowSeriesName = false;
                labels.ShowCategoryName = false;
                labels.ShowValue = true;
                labels.NumberFormat = "#,##0.##";
                labels.Font.Name = "微软雅黑";
                labels.Font.Size = 9;
                labels.Font.Color = ToOle("#34473F");
                var normalized = (chartType ?? string.Empty).ToLowerInvariant();
                labels.Position = normalized == "line"
                    ? XlDataLabelPosition.xlLabelPositionAbove
                    : XlDataLabelPosition.xlLabelPositionOutsideEnd;
            });
        }

        private static void StyleAxes(Microsoft.Office.Interop.Excel.Chart chart, SeriesData data)
        {
            Try(() =>
            {
                dynamic categoryAxis = chart.Axes(XlAxisType.xlCategory, XlAxisGroup.xlPrimary);
                categoryAxis.TickLabels.Font.Name = "微软雅黑";
                categoryAxis.TickLabels.Font.Size = 10;
                categoryAxis.TickLabels.Font.Color = ToOle("#53645F");
                categoryAxis.Format.Line.ForeColor.RGB = ToOle("#DCE6E2");
                if (data != null && data.Categories.Count > 12)
                    categoryAxis.TickLabelSpacing = (int)Math.Ceiling(data.Categories.Count / 12d);
            });
            Try(() =>
            {
                dynamic valueAxis = chart.Axes(XlAxisType.xlValue, XlAxisGroup.xlPrimary);
                valueAxis.TickLabels.Font.Name = "微软雅黑";
                valueAxis.TickLabels.Font.Size = 10;
                valueAxis.TickLabels.Font.Color = ToOle("#53645F");
                valueAxis.Format.Line.Visible = 0;
                valueAxis.MajorGridlines.Format.Line.ForeColor.RGB = ToOle("#E9EFEC");
            });
            Try(() =>
            {
                if (data == null || data.Values.Count == 0) return;
                var unit = DetermineDisplayUnit(data.Values);
                if (!unit.HasValue) return;
                var valueAxis = (Axis)chart.Axes(XlAxisType.xlValue, XlAxisGroup.xlPrimary);
                valueAxis.DisplayUnit = unit.Value;
                valueAxis.HasDisplayUnitLabel = true;
                valueAxis.DisplayUnitLabel.Caption = unit.Value == XlDisplayUnit.xlThousands
                    ? "K"
                    : unit.Value == XlDisplayUnit.xlMillions ? "M" : "B";
            });
        }

        private void ApplyLegend(Microsoft.Office.Interop.Excel.Chart chart, string position)
        {
            var normalized = (position ?? "bottom").Trim().ToLowerInvariant();
            if (normalized == "none")
            {
                chart.HasLegend = false;
                return;
            }
            chart.HasLegend = true;
            chart.Legend.Position = normalized == "right"
                ? XlLegendPosition.xlLegendPositionRight
                : XlLegendPosition.xlLegendPositionBottom;
            Try(() =>
            {
                chart.Legend.Font.Name = "微软雅黑";
                chart.Legend.Font.Size = 10;
                chart.Legend.Font.Color = ToOle("#53645F");
            });
        }

        private void AddCenterSummary(Microsoft.Office.Interop.Excel.Chart chart, SeriesData data)
        {
            Try(() =>
            {
                var total = 0d;
                for (var i = 0; i < data.Values.Count; i++) total += data.Values[i];
                const double width = 112;
                const double height = 56;
                var left = (_width - width) / 2;
                var top = (_height - height) / 2 + 14;
                dynamic shape = chart.Shapes.AddTextbox(
                    Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,
                    (float)left, (float)top, (float)width, (float)height);
                shape.Left = (float)left;
                shape.Top = (float)top;
                shape.Width = (float)width;
                shape.Height = (float)height;
                shape.TextFrame2.TextRange.Text = "合计\n" + FormatCompactNumber(total);
                shape.TextFrame2.TextRange.ParagraphFormat.Alignment = 2;
                shape.TextFrame2.VerticalAnchor = 3;
                shape.TextFrame2.TextRange.Font.Name = "微软雅黑";
                shape.TextFrame2.TextRange.Font.Size = 12;
                shape.TextFrame2.TextRange.Font.Bold = -1;
                shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = ToOle("#18352D");
                shape.Fill.Visible = 0;
                shape.Line.Visible = 0;
                shape.ZOrder(Microsoft.Office.Core.MsoZOrderCmd.msoBringToFront);
            });
        }

        private static bool TrySetOutsideLabels(Series series, int count)
        {
            var positioned = 0;
            for (var index = 1; index <= count; index++)
            {
                try
                {
                    dynamic label = series.DataLabels(index);
                    label.Position = XlDataLabelPosition.xlLabelPositionOutsideEnd;
                    if (Convert.ToInt32(label.Position) == (int)XlDataLabelPosition.xlLabelPositionOutsideEnd)
                        positioned++;
                }
                catch { }
            }
            return positioned == count;
        }

        private static string GetContrastTextColor(string htmlColor)
        {
            var color = ColorTranslator.FromHtml(htmlColor);
            var luminance = (0.2126 * color.R + 0.7152 * color.G + 0.0722 * color.B) / 255d;
            return luminance < 0.56 ? "#FFFFFF" : "#18352D";
        }

        private static string FormatCompactNumber(double value)
        {
            var absolute = Math.Abs(value);
            if (absolute >= 100000000) return (value / 100000000d).ToString("0.#", CultureInfo.InvariantCulture) + "亿";
            if (absolute >= 10000) return (value / 10000d).ToString("0.#", CultureInfo.InvariantCulture) + "万";
            return value.ToString("#,##0.##", CultureInfo.InvariantCulture);
        }

        private static string[] GetPalette(string name)
        {
            switch ((name ?? "emerald").Trim().ToLowerInvariant())
            {
                case "ocean": return new[] { "#164E63", "#0E7490", "#0891B2", "#22D3EE", "#67E8F9", "#A5F3FC", "#D97706" };
                case "sunset": return new[] { "#7C2D12", "#C2410C", "#EA580C", "#F59E0B", "#FBBF24", "#FCD34D", "#9A3412" };
                case "vivid": return new[] { "#5B3FD6", "#0EA5A6", "#F97316", "#E43D69", "#2F80ED", "#8B5CF6", "#16A34A" };
                default: return new[] { "#0F5B4B", "#1B7F67", "#36A486", "#F0B35A", "#E97862", "#5C8FD6", "#9A7BCB" };
            }
        }

        private static int ToOle(string htmlColor)
        {
            return ColorTranslator.ToOle(ColorTranslator.FromHtml(htmlColor));
        }

        private static void Try(System.Action action)
        {
            try { action(); } catch { }
        }

        private static string MakeUniqueName(ChartObjects chartObjects, string preferredName)
        {
            var baseName = string.IsNullOrWhiteSpace(preferredName) ? "AgentChart" : preferredName.Trim();
            var candidate = baseName;
            var suffix = 2;
            while (ContainsName(chartObjects, candidate)) candidate = baseName + suffix++;
            return candidate;
        }

        private static bool ContainsName(ChartObjects chartObjects, string name)
        {
            for (var i = 1; i <= chartObjects.Count; i++)
                if (string.Equals(((ChartObject)chartObjects.Item(i)).Name, name, StringComparison.OrdinalIgnoreCase))
                    return true;
            return false;
        }

        private static bool IsProportionChart(string value)
        {
            var normalized = (value ?? "").Trim().ToLowerInvariant();
            return normalized == "pie" || normalized == "doughnut";
        }

        private static XlChartType ParseChartType(string value)
        {
            switch ((value ?? "").Trim().ToLowerInvariant())
            {
                case "column": return XlChartType.xlColumnClustered;
                case "bar": return XlChartType.xlBarClustered;
                case "line": return XlChartType.xlLineMarkers;
                case "pie": return XlChartType.xlPie;
                case "doughnut": return XlChartType.xlDoughnut;
                case "area": return XlChartType.xlArea;
                case "scatter": return XlChartType.xlXYScatter;
                default: throw new ArgumentException("chart_type 仅支持 column、bar、line、pie、doughnut、area、scatter。");
            }
        }

        private sealed class SeriesData
        {
            public string CategoryHeader { get; set; }
            public string ValueHeader { get; set; }
            public int OriginalCount { get; set; }
            public string ProcessingNote { get; set; }
            public List<string> Categories { get; } = new List<string>();
            public List<double> Values { get; } = new List<double>();
        }

        private sealed class AggregateBucket
        {
            public string Label { get; set; }
            public double Sum { get; set; }
            public int Count { get; set; }
            public double Min { get; set; }
            public double Max { get; set; }
        }

        private sealed class ChronologicalPoint
        {
            public string Category { get; set; }
            public double Value { get; set; }
            public long Key { get; set; }
            public int Order { get; set; }
        }

        public sealed class Factory : IOperationFactory
        {
            public string ToolName => "chart_create";

            public IOperation Parse(string argumentsJson)
            {
                using (var doc = JsonDocument.Parse(argumentsJson))
                {
                    var root = doc.RootElement;
                    var sourceAddress = ReadString(root, "source_address");
                    var chartType = ReadString(root, "chart_type");
                    var title = ReadString(root, "title") ?? "数据图表";
                    var width = ReadDouble(root, "width", 600);
                    var height = ReadDouble(root, "height", 380);
                    var isProportion = IsProportionChart(chartType);
                    var aggregation = (ReadString(root, "aggregation") ?? "auto").Trim().ToLowerInvariant();
                    var maxCategories = ReadInt(root, "max_categories", isProportion ? 8 : 12);
                    var showLabelsByDefault = isProportion ||
                                              string.Equals(chartType, "column", StringComparison.OrdinalIgnoreCase) ||
                                              string.Equals(chartType, "bar", StringComparison.OrdinalIgnoreCase);
                    var legendPosition = (ReadString(root, "legend_position") ?? (isProportion ? "none" : "bottom")).ToLowerInvariant();
                    var palette = (ReadString(root, "palette") ?? "emerald").ToLowerInvariant();

                    if (string.IsNullOrWhiteSpace(sourceAddress)) throw new ArgumentException("source_address 不能为空。");
                    ParseChartType(chartType);
                    if (width < 320 || width > 1200) throw new ArgumentException("width 必须在 320 到 1200 之间。");
                    if (height < 220 || height > 900) throw new ArgumentException("height 必须在 220 到 900 之间。");
                    if (legendPosition != "none" && legendPosition != "bottom" && legendPosition != "right")
                        throw new ArgumentException("legend_position 仅支持 none、bottom 或 right。");
                    if (palette != "emerald" && palette != "ocean" && palette != "sunset" && palette != "vivid")
                        throw new ArgumentException("palette 仅支持 emerald、ocean、sunset 或 vivid。");
                    if (aggregation != "auto" && aggregation != "none" && aggregation != "sum" &&
                        aggregation != "average" && aggregation != "count" && aggregation != "min" && aggregation != "max")
                        throw new ArgumentException("aggregation 仅支持 auto、none、sum、average、count、min 或 max。");
                    if (maxCategories < 3 || maxCategories > 50)
                        throw new ArgumentException("max_categories 必须在 3 到 50 之间。");

                    return new CreateChartOp(
                        ReadString(root, "source_sheet"), sourceAddress,
                        ReadString(root, "destination_sheet"), ReadString(root, "anchor_address"),
                        chartType, title, ReadString(root, "name"), width, height,
                        ReadString(root, "category_field"), ReadString(root, "value_field"),
                        ReadStringSet(root, "exclude_categories"),
                        ReadBool(root, "sort_descending", isProportion),
                        ReadBool(root, "show_data_labels", showLabelsByDefault),
                        ReadBool(root, "show_percentage", isProportion),
                        legendPosition, palette, aggregation, maxCategories,
                        ReadBool(root, "include_other", isProportion));
                }
            }

            private static string ReadString(JsonElement root, string name)
            {
                return root.TryGetProperty(name, out var value) && value.ValueKind == JsonValueKind.String
                    ? value.GetString()
                    : null;
            }

            private static double ReadDouble(JsonElement root, string name, double fallback)
            {
                if (!root.TryGetProperty(name, out var value)) return fallback;
                if (value.ValueKind != JsonValueKind.Number) throw new ArgumentException(name + " 必须是数字。");
                return value.GetDouble();
            }

            private static int ReadInt(JsonElement root, string name, int fallback)
            {
                if (!root.TryGetProperty(name, out var value)) return fallback;
                if (value.ValueKind != JsonValueKind.Number || !value.TryGetInt32(out var result))
                    throw new ArgumentException(name + " 必须是整数。");
                return result;
            }

            private static bool ReadBool(JsonElement root, string name, bool fallback)
            {
                if (!root.TryGetProperty(name, out var value)) return fallback;
                if (value.ValueKind != JsonValueKind.True && value.ValueKind != JsonValueKind.False)
                    throw new ArgumentException(name + " 必须是布尔值。");
                return value.GetBoolean();
            }

            private static HashSet<string> ReadStringSet(JsonElement root, string name)
            {
                var result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                if (!root.TryGetProperty(name, out var value)) return result;
                if (value.ValueKind != JsonValueKind.Array) throw new ArgumentException(name + " 必须是字符串数组。");
                foreach (var item in value.EnumerateArray())
                {
                    if (item.ValueKind != JsonValueKind.String) throw new ArgumentException(name + " 必须是字符串数组。");
                    var text = item.GetString()?.Trim();
                    if (!string.IsNullOrWhiteSpace(text)) result.Add(text);
                }
                return result;
            }
        }
    }
}
