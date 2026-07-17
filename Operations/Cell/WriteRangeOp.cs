using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.Json;
using Microsoft.Office.Interop.Excel;

namespace AgentForExcel.Operations.Cell
{
    /// <summary>向目标区域批量写入普通值，不接受公式。</summary>
    public sealed class WriteRangeOp : IOperation
    {
        public string ToolName => "cell_write_range";
        public bool IsWriteOperation => true;

        private readonly string _sheetName;
        private readonly string _address;
        private readonly object[,] _values;

        private WriteRangeOp(string sheetName, string address, object[,] values)
        {
            _sheetName = sheetName;
            _address = address;
            _values = values;
        }

        public string Describe()
        {
            var where = string.IsNullOrWhiteSpace(_sheetName) ? "活动工作表" : "工作表「" + _sheetName + "」";
            return $"向 {where} 的 {_address} 写入 {_values.GetLength(0)} 行 × {_values.GetLength(1)} 列普通值";
        }

        public string Execute(AppContext context)
        {
            var sheet = CellOperationSupport.GetWorksheet(context, _sheetName);
            var target = CellOperationSupport.GetRange(sheet, _address);
            var rows = _values.GetLength(0);
            var columns = _values.GetLength(1);

            if (target.Rows.Count == 1 && target.Columns.Count == 1 && (rows > 1 || columns > 1))
                target = target.get_Resize(rows, columns);

            CellOperationSupport.ValidateRangeSize(target);
            if (target.Rows.Count != rows || target.Columns.Count != columns)
                throw new ArgumentException(
                    $"目标区域尺寸为 {target.Rows.Count} 行 × {target.Columns.Count} 列，但 values 是 {rows} 行 × {columns} 列。");

            target.Value2 = _values;
            return $"已写入 {sheet.Name}!{target.Address}，共 {rows * columns} 个单元格。";
        }

        public sealed class Factory : IOperationFactory
        {
            public string ToolName => "cell_write_range";

            public IOperation Parse(string argumentsJson)
            {
                using (var doc = JsonDocument.Parse(argumentsJson))
                {
                    var root = doc.RootElement;
                    var sheet = ReadString(root, "sheet");
                    var address = ReadString(root, "address");
                    if (string.IsNullOrWhiteSpace(address))
                        throw new ArgumentException("address 不能为空。");
                    if (!root.TryGetProperty("values", out var valuesElement))
                        throw new ArgumentException("values 不能为空。");

                    return new WriteRangeOp(sheet, address, ParseValues(valuesElement));
                }
            }

            private static object[,] ParseValues(JsonElement element)
            {
                if (element.ValueKind != JsonValueKind.Array)
                    return new[,] { { ParseScalar(element) } };

                var outer = new List<JsonElement>();
                foreach (var item in element.EnumerateArray()) outer.Add(item);
                if (outer.Count == 0) throw new ArgumentException("values 不能为空数组。");

                if (outer[0].ValueKind != JsonValueKind.Array)
                {
                    var row = new object[1, outer.Count];
                    for (var c = 0; c < outer.Count; c++) row[0, c] = ParseScalar(outer[c]);
                    return row;
                }

                var rows = new List<List<object>>();
                var columnCount = -1;
                foreach (var rowElement in outer)
                {
                    if (rowElement.ValueKind != JsonValueKind.Array)
                        throw new ArgumentException("values 必须全部是一维行数组，不能混合标量和数组。");

                    var row = new List<object>();
                    foreach (var cell in rowElement.EnumerateArray()) row.Add(ParseScalar(cell));
                    if (row.Count == 0) throw new ArgumentException("values 中不能包含空行。");
                    if (columnCount < 0) columnCount = row.Count;
                    if (row.Count != columnCount) throw new ArgumentException("values 每一行的列数必须一致。");
                    rows.Add(row);
                }

                var matrix = new object[rows.Count, columnCount];
                for (var r = 0; r < rows.Count; r++)
                    for (var c = 0; c < columnCount; c++)
                        matrix[r, c] = rows[r][c];
                return matrix;
            }

            private static object ParseScalar(JsonElement element)
            {
                switch (element.ValueKind)
                {
                    case JsonValueKind.Null:
                        return null;
                    case JsonValueKind.String:
                        var text = element.GetString();
                        if (!string.IsNullOrEmpty(text) && text.TrimStart().StartsWith("=", StringComparison.Ordinal))
                            throw new InvalidOperationException("普通值不能以 = 开头；写入公式请使用 cell_fill_formula。");
                        return text;
                    case JsonValueKind.Number:
                        if (element.TryGetInt64(out var integer)) return integer;
                        return element.GetDouble();
                    case JsonValueKind.True:
                    case JsonValueKind.False:
                        return element.GetBoolean();
                    default:
                        throw new ArgumentException("values 仅支持字符串、数字、布尔值和 null。");
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
