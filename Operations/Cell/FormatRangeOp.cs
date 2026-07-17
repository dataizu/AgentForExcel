using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.Json;
using Microsoft.Office.Interop.Excel;

namespace AgentForExcel.Operations.Cell
{
    /// <summary>设置区域的常用显示格式。</summary>
    public sealed class FormatRangeOp : IOperation
    {
        public string ToolName => "cell_format_range";
        public bool IsWriteOperation => true;

        private readonly string _sheetName;
        private readonly string _address;
        private readonly string _numberFormat;
        private readonly bool? _bold;
        private readonly bool? _italic;
        private readonly double? _fontSize;
        private readonly string _fontColor;
        private readonly string _fillColor;
        private readonly string _horizontalAlignment;
        private readonly bool? _wrapText;
        private readonly bool? _addBorders;
        private readonly bool? _autofitColumns;
        private readonly bool? _autofitRows;

        private FormatRangeOp(
            string sheetName, string address, string numberFormat, bool? bold, bool? italic,
            double? fontSize, string fontColor, string fillColor, string horizontalAlignment,
            bool? wrapText, bool? addBorders, bool? autofitColumns, bool? autofitRows)
        {
            _sheetName = sheetName;
            _address = address;
            _numberFormat = numberFormat;
            _bold = bold;
            _italic = italic;
            _fontSize = fontSize;
            _fontColor = fontColor;
            _fillColor = fillColor;
            _horizontalAlignment = horizontalAlignment;
            _wrapText = wrapText;
            _addBorders = addBorders;
            _autofitColumns = autofitColumns;
            _autofitRows = autofitRows;
        }

        public string Describe()
        {
            var where = string.IsNullOrWhiteSpace(_sheetName) ? "活动工作表" : "工作表「" + _sheetName + "」";
            return $"设置 {where} 的 {_address} 基础格式：{string.Join("、", GetChangeNames())}";
        }

        public string Execute(AppContext context)
        {
            var sheet = CellOperationSupport.GetWorksheet(context, _sheetName);
            var target = CellOperationSupport.GetRange(sheet, _address);

            if (_numberFormat != null) target.NumberFormat = _numberFormat;
            if (_bold.HasValue) target.Font.Bold = _bold.Value;
            if (_italic.HasValue) target.Font.Italic = _italic.Value;
            if (_fontSize.HasValue) target.Font.Size = _fontSize.Value;
            if (_fontColor != null) target.Font.Color = CellOperationSupport.ToOleColor(_fontColor, "font_color");
            if (_fillColor != null) target.Interior.Color = CellOperationSupport.ToOleColor(_fillColor, "fill_color");
            if (_wrapText.HasValue) target.WrapText = _wrapText.Value;
            if (_horizontalAlignment != null) target.HorizontalAlignment = ParseAlignment(_horizontalAlignment);

            if (_addBorders == true)
            {
                target.Borders.LineStyle = XlLineStyle.xlContinuous;
                target.Borders.Weight = XlBorderWeight.xlThin;
                target.Borders.Color = CellOperationSupport.ToOleColor("#D8E1DA", "border_color");
            }
            if (_autofitColumns == true) target.EntireColumn.AutoFit();
            if (_autofitRows == true) target.EntireRow.AutoFit();

            return $"已设置 {sheet.Name}!{target.Address} 的格式：{string.Join("、", GetChangeNames())}。";
        }

        private List<string> GetChangeNames()
        {
            var changes = new List<string>();
            if (_numberFormat != null) changes.Add("数字格式");
            if (_bold.HasValue) changes.Add(_bold.Value ? "加粗" : "取消加粗");
            if (_italic.HasValue) changes.Add(_italic.Value ? "斜体" : "取消斜体");
            if (_fontSize.HasValue) changes.Add("字号");
            if (_fontColor != null) changes.Add("字体颜色");
            if (_fillColor != null) changes.Add("填充颜色");
            if (_horizontalAlignment != null) changes.Add("水平对齐");
            if (_wrapText.HasValue) changes.Add("自动换行");
            if (_addBorders == true) changes.Add("边框");
            if (_autofitColumns == true) changes.Add("自适应列宽");
            if (_autofitRows == true) changes.Add("自适应行高");
            return changes;
        }

        private static XlHAlign ParseAlignment(string value)
        {
            switch ((value ?? "").Trim().ToLowerInvariant())
            {
                case "left": return XlHAlign.xlHAlignLeft;
                case "center": return XlHAlign.xlHAlignCenter;
                case "right": return XlHAlign.xlHAlignRight;
                case "general": return XlHAlign.xlHAlignGeneral;
                default: throw new ArgumentException("horizontal_alignment 仅支持 left、center、right、general。");
            }
        }

        public sealed class Factory : IOperationFactory
        {
            public string ToolName => "cell_format_range";

            public IOperation Parse(string argumentsJson)
            {
                using (var doc = JsonDocument.Parse(argumentsJson))
                {
                    var root = doc.RootElement;
                    var address = ReadString(root, "address");
                    if (string.IsNullOrWhiteSpace(address)) throw new ArgumentException("address 不能为空。");

                    var op = new FormatRangeOp(
                        ReadString(root, "sheet"), address, ReadString(root, "number_format"),
                        ReadNullableBool(root, "bold"), ReadNullableBool(root, "italic"),
                        ReadNullableDouble(root, "font_size"), ReadString(root, "font_color"),
                        ReadString(root, "fill_color"), ReadString(root, "horizontal_alignment"),
                        ReadNullableBool(root, "wrap_text"), ReadNullableBool(root, "add_borders"),
                        ReadNullableBool(root, "autofit_columns"), ReadNullableBool(root, "autofit_rows"));

                    if (op.GetChangeNames().Count == 0)
                        throw new ArgumentException("至少需要提供一个格式参数。");
                    if (op._fontSize.HasValue && (op._fontSize < 6 || op._fontSize > 72))
                        throw new ArgumentException("font_size 必须在 6 到 72 之间。");
                    if (op._fontColor != null) CellOperationSupport.ToOleColor(op._fontColor, "font_color");
                    if (op._fillColor != null) CellOperationSupport.ToOleColor(op._fillColor, "fill_color");
                    if (op._horizontalAlignment != null) ParseAlignment(op._horizontalAlignment);
                    return op;
                }
            }

            private static string ReadString(JsonElement root, string name)
            {
                return root.TryGetProperty(name, out var value) && value.ValueKind == JsonValueKind.String
                    ? value.GetString()
                    : null;
            }

            private static bool? ReadNullableBool(JsonElement root, string name)
            {
                if (!root.TryGetProperty(name, out var value)) return null;
                if (value.ValueKind != JsonValueKind.True && value.ValueKind != JsonValueKind.False)
                    throw new ArgumentException(name + " 必须是 true 或 false。");
                return value.GetBoolean();
            }

            private static double? ReadNullableDouble(JsonElement root, string name)
            {
                if (!root.TryGetProperty(name, out var value)) return null;
                if (value.ValueKind != JsonValueKind.Number)
                    throw new ArgumentException(name + " 必须是数字。");
                return value.GetDouble();
            }
        }
    }
}
