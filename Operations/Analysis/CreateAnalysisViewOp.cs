using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.Json;
using Microsoft.Office.Interop.Excel;

namespace AgentForExcel.Operations.Analysis
{
    /// <summary>把源数据复制为值快照，在新的分析工作表中安全排序和展示。</summary>
    public sealed class CreateAnalysisViewOp : IOperation
    {
        public string ToolName => "analysis_create_view";
        public bool IsWriteOperation => true;

        private readonly string _sourceSheet;
        private readonly string _sourceAddress;
        private readonly string _analysisSheetName;
        private readonly string _destinationAddress;
        private readonly IReadOnlyList<SortField> _sortFields;

        private CreateAnalysisViewOp(
            string sourceSheet,
            string sourceAddress,
            string analysisSheetName,
            string destinationAddress,
            IReadOnlyList<SortField> sortFields)
        {
            _sourceSheet = sourceSheet;
            _sourceAddress = sourceAddress;
            _analysisSheetName = analysisSheetName;
            _destinationAddress = destinationAddress;
            _sortFields = sortFields ?? new List<SortField>();
        }

        public string Describe()
        {
            var sort = _sortFields.Count == 0 ? "保留原顺序" : "按指定字段排序";
            return $"把 {_sourceSheet ?? "活动工作表"}!{_sourceAddress} 复制到新的分析工作表并{sort}；原始数据不会被修改";
        }

        public string Execute(AppContext context)
        {
            // 分析页会整块写入快照并做大量格式化,批量作用域抑制逐次重绘。
            using (new ExcelBatchScope(context))
            {
                return ExecuteCore(context);
            }
        }

        private string ExecuteCore(AppContext context)
        {
            var sourceSheet = Cell.CellOperationSupport.GetWorksheet(context, _sourceSheet);
            var sourceRange = Cell.CellOperationSupport.GetRange(sourceSheet, _sourceAddress);
            if (sourceRange.Rows.Count < 2)
                throw new ArgumentException("源区域至少需要标题行和一行数据。");

            var headers = ReadHeaders(sourceRange);
            var analysisSheet = AnalysisSheetSupport.CreateUniqueWorksheet(context, _analysisSheetName);
            try
            {
                var topLeft = Cell.CellOperationSupport.GetRange(analysisSheet, _destinationAddress);
                var destination = analysisSheet.Range[
                    topLeft,
                    analysisSheet.Cells[topLeft.Row + sourceRange.Rows.Count - 1, topLeft.Column + sourceRange.Columns.Count - 1]];

                // 只复制当前值，源表公式不会带入分析页，确保分析操作不反向影响源数据。
                var snapshot = (object[,])sourceRange.Value2;
                if (_sortFields.Count > 0)
                    snapshot = SortSnapshot(snapshot, headers, _sortFields);
                destination.Value2 = snapshot;
                try { destination.NumberFormat = sourceRange.NumberFormat; } catch { }

                var table = analysisSheet.ListObjects.Add(
                    XlListObjectSourceType.xlSrcRange,
                    destination,
                    Type.Missing,
                    XlYesNoGuess.xlYes,
                    Type.Missing);
                table.Name = "AgentView" + DateTime.Now.Ticks.ToString().Substring(8);
                table.TableStyle = "TableStyleMedium4";

                destination.Columns.AutoFit();
                for (var column = 1; column <= destination.Columns.Count; column++)
                {
                    var targetColumn = (Range)destination.Columns[column];
                    if (Convert.ToDouble(targetColumn.ColumnWidth) > 28) targetColumn.ColumnWidth = 28;
                }
                ((Range)destination.Rows[1]).RowHeight = 24;

                return $"已创建安全分析视图「{analysisSheet.Name}」，数据位于 {destination.Address}；原始工作表「{sourceSheet.Name}」未修改。";
            }
            catch
            {
                AnalysisSheetSupport.DeleteWorksheetSilently(context, analysisSheet);
                throw;
            }
        }

        private static Dictionary<string, int> ReadHeaders(Range sourceRange)
        {
            var result = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (var column = 1; column <= sourceRange.Columns.Count; column++)
            {
                var header = Convert.ToString(((Range)sourceRange.Cells[1, column]).Value2)?.Trim();
                if (string.IsNullOrWhiteSpace(header))
                    throw new ArgumentException($"源区域第 {column} 列标题为空。");
                if (result.ContainsKey(header))
                    throw new ArgumentException("源区域存在重复字段名：「" + header + "」。");
                result[header] = column;
            }
            return result;
        }

        private static object[,] SortSnapshot(
            object[,] snapshot,
            IReadOnlyDictionary<string, int> headers,
            IReadOnlyList<SortField> fields)
        {
            if (fields.Count > 3)
                throw new ArgumentException("一次最多支持 3 个排序字段。");

            var columns = new int[fields.Count];
            for (var index = 0; index < fields.Count; index++)
            {
                if (!headers.TryGetValue(fields[index].Field, out var column))
                    throw new ArgumentException("找不到排序字段：「" + fields[index].Field + "」。");
                columns[index] = column;
            }

            var firstRow = snapshot.GetLowerBound(0);
            var lastRow = snapshot.GetUpperBound(0);
            var firstColumn = snapshot.GetLowerBound(1);
            var lastColumn = snapshot.GetUpperBound(1);
            var rowIndexes = new List<int>();
            for (var row = firstRow + 1; row <= lastRow; row++) rowIndexes.Add(row);
            rowIndexes.Sort((leftRow, rightRow) =>
            {
                for (var index = 0; index < fields.Count; index++)
                {
                    var comparison = CompareValues(snapshot[leftRow, columns[index]], snapshot[rightRow, columns[index]]);
                    if (comparison == 0) continue;
                    return fields[index].Descending ? -comparison : comparison;
                }
                return leftRow.CompareTo(rightRow);
            });

            var sorted = new object[lastRow - firstRow + 1, lastColumn - firstColumn + 1];
            for (var column = firstColumn; column <= lastColumn; column++)
                sorted[0, column - firstColumn] = snapshot[firstRow, column];
            for (var targetRow = 1; targetRow < sorted.GetLength(0); targetRow++)
            {
                var sourceRow = rowIndexes[targetRow - 1];
                for (var column = firstColumn; column <= lastColumn; column++)
                    sorted[targetRow, column - firstColumn] = snapshot[sourceRow, column];
            }
            return sorted;
        }

        private static int CompareValues(object left, object right)
        {
            if (left == null && right == null) return 0;
            if (left == null) return 1;
            if (right == null) return -1;

            double leftNumber;
            double rightNumber;
            if (double.TryParse(Convert.ToString(left, CultureInfo.InvariantCulture), NumberStyles.Any, CultureInfo.InvariantCulture, out leftNumber) &&
                double.TryParse(Convert.ToString(right, CultureInfo.InvariantCulture), NumberStyles.Any, CultureInfo.InvariantCulture, out rightNumber))
                return leftNumber.CompareTo(rightNumber);

            return string.Compare(
                Convert.ToString(left, CultureInfo.CurrentCulture),
                Convert.ToString(right, CultureInfo.CurrentCulture),
                StringComparison.CurrentCultureIgnoreCase);
        }

        private sealed class SortField
        {
            public string Field { get; set; }
            public bool Descending { get; set; }
        }

        public sealed class Factory : IOperationFactory
        {
            public string ToolName => "analysis_create_view";

            public IOperation Parse(string argumentsJson)
            {
                using (var doc = JsonDocument.Parse(argumentsJson))
                {
                    var root = doc.RootElement;
                    var sourceAddress = ReadString(root, "source_address");
                    if (string.IsNullOrWhiteSpace(sourceAddress))
                        throw new ArgumentException("source_address 不能为空。");

                    var sortFields = new List<SortField>();
                    if (root.TryGetProperty("sort_by", out var sortBy) && sortBy.ValueKind == JsonValueKind.Array)
                    {
                        foreach (var item in sortBy.EnumerateArray())
                        {
                            var field = ReadString(item, "field");
                            if (string.IsNullOrWhiteSpace(field)) throw new ArgumentException("sort_by.field 不能为空。");
                            var direction = (ReadString(item, "direction") ?? "asc").ToLowerInvariant();
                            if (direction != "asc" && direction != "desc")
                                throw new ArgumentException("sort_by.direction 仅支持 asc 或 desc。");
                            sortFields.Add(new SortField { Field = field, Descending = direction == "desc" });
                        }
                    }

                    return new CreateAnalysisViewOp(
                        ReadString(root, "source_sheet"),
                        sourceAddress,
                        ReadString(root, "analysis_sheet_name") ?? "Agent分析",
                        ReadString(root, "destination_address") ?? "A1",
                        sortFields);
                }
            }

            private static string ReadString(JsonElement root, string name)
            {
                return root.TryGetProperty(name, out var value) && value.ValueKind == JsonValueKind.String
                    ? value.GetString()
                    : null;
            }
        }
    }
}
