using System;
using System.Collections.Generic;
using System.Text.Json;
using Microsoft.Office.Interop.Excel;

namespace AgentForExcel.Operations.Cell
{
    /// <summary>
    /// 一次性绘制像素画：传入 #RRGGBB 二维颜色矩阵，自动设置列宽/行高，
    /// 并把每行水平同色连续段合并为单个矩形填色，一次调用完成整幅图。
    /// 空字符串("")或 null 表示跳过、保持空白。
    /// pixel_width / pixel_height 单位为磅(pt)，默认 12×12，保证每个像素格接近正方形；
    /// hide_gridlines 默认 true，隐藏网格线让像素画更清晰。
    /// </summary>
    public sealed class DrawPixelsOp : IOperation
    {
        public string ToolName => "cell_draw_pixels";
        public bool IsWriteOperation => true;

        private readonly string _sheetName;
        private readonly string _address;
        private readonly string[,] _pixels;
        private readonly double _pixelWidth;
        private readonly double _pixelHeight;
        private readonly bool _hideGridlines;

        internal DrawPixelsOp(string sheetName, string address, string[,] pixels, double pixelWidth, double pixelHeight, bool hideGridlines)
        {
            _sheetName = sheetName;
            _address = address;
            _pixels = pixels;
            _pixelWidth = pixelWidth;
            _pixelHeight = pixelHeight;
            _hideGridlines = hideGridlines;
        }

        public string Describe()
        {
            var where = string.IsNullOrWhiteSpace(_sheetName) ? "活动工作表" : "工作表「" + _sheetName + "」";
            return $"在 {where} 的 {_address} 绘制 {_pixels.GetLength(0)} 行 × {_pixels.GetLength(1)} 列的像素画";
        }

        public string Execute(AppContext context)
        {
            // 像素绘制会产生上万次 Interior 赋值,批量作用域避免逐段重绘与事件链。
            using (new ExcelBatchScope(context))
            {
                return ExecuteCore(context);
            }
        }

        private string ExecuteCore(AppContext context)
        {
            var sheet = CellOperationSupport.GetWorksheet(context, _sheetName);
            var rows = _pixels.GetLength(0);
            var columns = _pixels.GetLength(1);

            // 只取左上角单元格，再扩展到像素矩阵尺寸
            var start = CellOperationSupport.GetRange(sheet, _address).Cells[1, 1] as Range;
            if (start == null) throw new ArgumentException("无法解析左上角单元格：" + _address);
            var target = start.get_Resize(rows, columns);
            CellOperationSupport.ValidateRangeSize(target);

            SetSquarePixels(target, start, _pixelWidth, _pixelHeight);

            // 预解析颜色缓存，避免逐格重复转换
            var colorCache = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (var r = 0; r < rows; r++)
            {
                for (var c = 0; c < columns; c++)
                {
                    var pixel = _pixels[r, c];
                    if (string.IsNullOrEmpty(pixel)) continue;
                    if (!colorCache.ContainsKey(pixel))
                        colorCache[pixel] = CellOperationSupport.ToOleColor(pixel, "pixels");
                }
            }

            // 每行水平同色连续段合并为一个矩形，一次 COM 调用填色
            var segmentCount = 0;
            var coloredCells = 0;
            for (var r = 0; r < rows; r++)
            {
                var c = 0;
                while (c < columns)
                {
                    var pixel = _pixels[r, c];
                    if (string.IsNullOrEmpty(pixel)) { c++; continue; }

                    var end = c;
                    while (end + 1 < columns &&
                           !string.IsNullOrEmpty(_pixels[r, end + 1]) &&
                           string.Equals(_pixels[r, end + 1], pixel, StringComparison.OrdinalIgnoreCase))
                        end++;

                    var segment = start.get_Offset(r, c).get_Resize(1, end - c + 1);
                    segment.Interior.Color = colorCache[pixel];
                    segmentCount++;
                    coloredCells += end - c + 1;
                    c = end + 1;
                }
            }

            if (_hideGridlines)
            {
                try
                {
                    var window = sheet.Application.ActiveWindow as Window;
                    if (window != null) window.DisplayGridlines = false;
                }
                catch
                {
                    // 隐藏网格线只是视觉优化，失败不影响绘制结果
                }
            }

            return $"已在 {sheet.Name}!{target.Address} 绘制像素画：{rows} 行 × {columns} 列，共 {coloredCells} 个着色单元格（合并为 {segmentCount} 个矩形段），像素格约 {_pixelWidth:0.#}×{_pixelHeight:0.#} 磅。";
        }

        /// <summary>
        /// 让像素格接近正方形且大小可控：行高直接按磅设置；
        /// 列宽用「磅 → 字符宽度」初始换算后再按实测宽度迭代微调（target 负责整片行列尺寸，probe 单格测量），
        /// 不依赖具体字体，兼容不同 DPI 与默认字体。
        /// </summary>
        private static void SetSquarePixels(Range target, Range probe, double widthPt, double heightPt)
        {
            target.EntireRow.RowHeight = heightPt;

            // 初始换算：默认字体下 1 字符宽 ≈ 7 磅（含单元格边距），再迭代校正
            var columnWidth = (widthPt - 5.0) / 7.0;
            if (columnWidth < 0.5) columnWidth = 0.5;
            target.EntireColumn.ColumnWidth = columnWidth;

            for (var i = 0; i < 6; i++)
            {
                double actualWidth;
                try { actualWidth = Convert.ToDouble(probe.Width); }
                catch { break; }

                var error = actualWidth - widthPt;
                if (Math.Abs(error) < 0.25) break;

                columnWidth -= error / 7.0;
                if (columnWidth < 0.5) columnWidth = 0.5;
                if (columnWidth > 255) columnWidth = 255;
                target.EntireColumn.ColumnWidth = columnWidth;
            }
        }

        public sealed class Factory : IOperationFactory
        {
            public string ToolName => "cell_draw_pixels";

            public IOperation Parse(string argumentsJson)
            {
                using (var doc = JsonDocument.Parse(argumentsJson))
                {
                    var root = doc.RootElement;
                    var address = ReadString(root, "address");
                    if (string.IsNullOrWhiteSpace(address))
                        throw new ArgumentException("address 不能为空。");
                    if (!root.TryGetProperty("pixels", out var pixelsElement))
                        throw new ArgumentException("pixels 不能为空。");

                    var pixels = ParsePixels(pixelsElement);
                    var rows = pixels.GetLength(0);
                    var columns = pixels.GetLength(1);
                    var total = (long)rows * columns;
                    if (total > CellOperationSupport.MaxCellsPerOperation)
                        throw new InvalidOperationException(
                            $"像素矩阵 {rows} 行 × {columns} 列共 {total} 个单元格，超过单次上限 {CellOperationSupport.MaxCellsPerOperation:0}。");

                    var pixelWidth = ReadNullableDouble(root, "pixel_width") ?? 12.0;
                    var pixelHeight = ReadNullableDouble(root, "pixel_height") ?? 12.0;
                    if (pixelWidth < 1 || pixelWidth > 60)
                        throw new ArgumentException("pixel_width(磅) 必须在 1 到 60 之间。");
                    if (pixelHeight < 2 || pixelHeight > 100)
                        throw new ArgumentException("pixel_height(磅) 必须在 2 到 100 之间。");

                    var hideGridlines = ReadNullableBool(root, "hide_gridlines") ?? true;

                    return new DrawPixelsOp(ReadString(root, "sheet"), address, pixels, pixelWidth, pixelHeight, hideGridlines);
                }
            }

            private static string[,] ParsePixels(JsonElement element)
            {
                if (element.ValueKind != JsonValueKind.Array || element.GetArrayLength() == 0)
                    throw new ArgumentException("pixels 必须是非空二维数组。");

                var outer = new List<JsonElement>();
                foreach (var item in element.EnumerateArray()) outer.Add(item);

                if (outer[0].ValueKind != JsonValueKind.Array)
                    throw new ArgumentException("pixels 必须是二维数组，例如 [[\"#E52521\",\"\"],[\"\",\"#1A1A1A\"]]。");

                var rows = new List<List<string>>();
                var columnCount = -1;
                foreach (var rowElement in outer)
                {
                    if (rowElement.ValueKind != JsonValueKind.Array)
                        throw new ArgumentException("pixels 每一行都必须是数组，不能混合标量和数组。");

                    var row = new List<string>();
                    foreach (var cell in rowElement.EnumerateArray()) row.Add(ParseColor(cell));
                    if (row.Count == 0) throw new ArgumentException("pixels 中不能包含空行。");
                    if (columnCount < 0) columnCount = row.Count;
                    if (row.Count != columnCount)
                        throw new ArgumentException("pixels 每一行的列数必须一致。");
                    rows.Add(row);
                }

                var matrix = new string[rows.Count, columnCount];
                for (var r = 0; r < rows.Count; r++)
                    for (var c = 0; c < columnCount; c++)
                        matrix[r, c] = rows[r][c];
                return matrix;
            }

            private static string ParseColor(JsonElement element)
            {
                switch (element.ValueKind)
                {
                    case JsonValueKind.Null:
                        return null;
                    case JsonValueKind.String:
                        var text = element.GetString();
                        if (string.IsNullOrEmpty(text)) return null; // 空串表示跳过
                        CellOperationSupport.ToOleColor(text, "pixels"); // 校验格式
                        return text;
                    default:
                        throw new ArgumentException("pixels 仅支持 #RRGGBB 字符串、空串或 null。");
                }
            }

            private static string ReadString(JsonElement root, string name)
            {
                return root.TryGetProperty(name, out var value) && value.ValueKind == JsonValueKind.String
                    ? value.GetString()
                    : null;
            }

            private static double? ReadNullableDouble(JsonElement root, string name)
            {
                if (!root.TryGetProperty(name, out var value)) return null;
                if (value.ValueKind != JsonValueKind.Number)
                    throw new ArgumentException(name + " 必须是数字。");
                return value.GetDouble();
            }

            private static bool? ReadNullableBool(JsonElement root, string name)
            {
                if (!root.TryGetProperty(name, out var value)) return null;
                if (value.ValueKind != JsonValueKind.True && value.ValueKind != JsonValueKind.False)
                    throw new ArgumentException(name + " 必须是布尔值。");
                return value.GetBoolean();
            }
        }
    }
}