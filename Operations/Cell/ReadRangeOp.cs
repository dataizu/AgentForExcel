using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Text.Json;
using AgentForExcel.Services;
using Microsoft.Office.Interop.Excel;

namespace AgentForExcel.Operations.Cell
{
    /// <summary>
    /// 读取指定区域的值。只读、无副作用 —— 阶段 1 联调用,
    /// 同时作为"第一个能力闭环"的最小可行样本。
    /// </summary>
    public class ReadRangeOp : IOperation
    {
        public const string TablePayloadPrefix = "__AGENT_TABLE_PREVIEW__";
        public string ToolName => "cell_read_range";
        public bool IsWriteOperation => false;

        private readonly string _sheetName;
        private readonly string _address;

        private ReadRangeOp(string sheetName, string address)
        {
            _sheetName = sheetName;
            _address = address;
        }

        public string Describe()
        {
            var where = string.IsNullOrEmpty(_sheetName) ? "活动工作表" : ("工作表「" + _sheetName + "」");
            return $"读取 {where} 的 {_address}";
        }

        public string Execute(AppContext context)
        {
            Worksheet sheet = string.IsNullOrEmpty(_sheetName)
                ? (Worksheet)context.Excel.ActiveSheet
                : (Worksheet)context.Excel.ActiveWorkbook.Worksheets[_sheetName];

            Range range = string.IsNullOrEmpty(_address)
                ? sheet.UsedRange
                : sheet.get_Range(_address);

            var totalRows = Convert.ToInt32(range.Rows.Count);
            var totalColumns = Convert.ToInt32(range.Columns.Count);
            var shownRows = Math.Min(totalRows, 12);
            var shownColumns = Math.Min(totalColumns, 8);
            // 聊天内只展示 12 x 8 预览。不要为了预览而把完整 UsedRange 复制进托管内存。
            var previewRange = range.get_Resize(shownRows, shownColumns);
            var readTimer = Stopwatch.StartNew();
            var raw = previewRange.Value2;
            readTimer.Stop();
            var matrix = raw as object[,];
            var rows = new List<List<string>>();

            for (var row = 1; row <= shownRows; row++)
            {
                var cells = new List<string>();
                for (var column = 1; column <= shownColumns; column++)
                {
                    object value;
                    if (matrix != null)
                        value = matrix[row, column];
                    else
                        value = row == 1 && column == 1 ? raw : null;
                    cells.Add(FormatCellValue(value));
                }
                rows.Add(cells);
            }

            var startColumn = Convert.ToInt32(range.Column);
            var headers = new List<string>();
            for (var column = 0; column < shownColumns; column++)
                headers.Add(ToColumnName(startColumn + column));

            var address = range.get_Address(false, false, XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
            var payload = new
            {
                kind = "table",
                title = "已读取数据",
                sheet = Convert.ToString(sheet.Name),
                address,
                start_row = Convert.ToInt32(range.Row),
                total_rows = totalRows,
                total_columns = totalColumns,
                shown_rows = shownRows,
                shown_columns = shownColumns,
                truncated = shownRows < totalRows || shownColumns < totalColumns,
                headers,
                rows
            };
            PerformanceLogger.Log(
                "range_read",
                readTimer.ElapsedMilliseconds,
                "rows=" + totalRows + "|columns=" + totalColumns +
                "|preview_rows=" + shownRows + "|preview_columns=" + shownColumns);
            return TablePayloadPrefix + JsonSerializer.Serialize(payload);
        }

        private static string FormatCellValue(object value)
        {
            if (value == null) return string.Empty;
            var text = Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
            text = text.Replace("\r", " ").Replace("\n", " ");
            return text.Length <= 120 ? text : text.Substring(0, 117) + "…";
        }

        private static string ToColumnName(int columnNumber)
        {
            var name = string.Empty;
            while (columnNumber > 0)
            {
                columnNumber--;
                name = (char)('A' + columnNumber % 26) + name;
                columnNumber /= 26;
            }
            return name;
        }

        /// <summary>工厂:把 LLM 下发的 arguments JSON 解析成 ReadRangeOp。</summary>
        public sealed class Factory : IOperationFactory
        {
            public string ToolName => "cell_read_range";

            public IOperation Parse(string argumentsJson)
            {
                string sheet = null, address = null;
                using (var doc = JsonDocument.Parse(argumentsJson))
                {
                    if (doc.RootElement.TryGetProperty("sheet", out var s) && s.ValueKind == JsonValueKind.String)
                        sheet = s.GetString();
                    if (doc.RootElement.TryGetProperty("address", out var a) && a.ValueKind == JsonValueKind.String)
                        address = a.GetString();
                }
                if (string.IsNullOrWhiteSpace(address))
                    throw new ArgumentException("address 不能为空");
                return new ReadRangeOp(sheet, address);
            }
        }
    }
}
