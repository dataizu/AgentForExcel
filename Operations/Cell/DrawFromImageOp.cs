using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Text.Json;

namespace AgentForExcel.Operations.Cell
{
    /// <summary>
    /// 从本地图片文件转换为像素画：读取图片 → 按 grid_width × grid_height 网格缩放采样为色值矩阵
    /// → 复用 cell_draw_pixels 的绘制管线一次性画入工作表。
    /// 默认最近邻采样（像素画清晰锐利、颜色精确）；interpolation=bilinear 时平滑过渡。
    /// 可选 palette 做最近色量化（像素画风格）；透明区域（alpha &lt; 128）输出空串跳过、保持空白。
    /// </summary>
    public sealed class DrawFromImageOp : IOperation
    {
        public string ToolName => "cell_draw_from_image";
        public bool IsWriteOperation => true;

        private readonly DrawPixelsOp _inner;
        private readonly string _imagePath;

        internal DrawFromImageOp(DrawPixelsOp inner, string imagePath)
        {
            _inner = inner ?? throw new ArgumentNullException(nameof(inner));
            _imagePath = imagePath;
        }

        public string Describe()
        {
            return _inner.Describe() + "（图片来源：" + Path.GetFileName(_imagePath) + "）";
        }

        public string Execute(AppContext context)
        {
            return _inner.Execute(context);
        }

        public sealed class Factory : IOperationFactory
        {
            public string ToolName => "cell_draw_from_image";

            private static readonly string[] AllowedExtensions = { ".png", ".jpg", ".jpeg", ".bmp", ".gif" };
            private const long MaxImageBytes = 50L * 1024 * 1024; // 50 MB

            public IOperation Parse(string argumentsJson)
            {
                using (var doc = JsonDocument.Parse(argumentsJson))
                {
                    var root = doc.RootElement;
                    var address = ReadString(root, "address");
                    if (string.IsNullOrWhiteSpace(address))
                        throw new ArgumentException("address 不能为空。");

                    var imagePath = ReadString(root, "image_path");
                    if (string.IsNullOrWhiteSpace(imagePath))
                        throw new ArgumentException("image_path 不能为空。");

                    var fullPath = Path.GetFullPath(imagePath.Trim());
                    if (!File.Exists(fullPath))
                        throw new ArgumentException("找不到图片文件：" + fullPath);
                    var extension = Path.GetExtension(fullPath).ToLowerInvariant();
                    if (Array.IndexOf(AllowedExtensions, extension) < 0)
                        throw new ArgumentException("仅支持本地图片文件：" + string.Join("、", AllowedExtensions));
                    var fileInfo = new FileInfo(fullPath);
                    if (fileInfo.Length > MaxImageBytes)
                        throw new ArgumentException($"图片文件过大（{fileInfo.Length / 1024.0 / 1024.0:0.#} MB），上限 50 MB。");

                    var gridWidth = ReadRequiredInt(root, "grid_width");
                    var gridHeight = ReadRequiredInt(root, "grid_height");
                    var total = (long)gridWidth * gridHeight;
                    if (total > CellOperationSupport.MaxCellsPerOperation)
                        throw new InvalidOperationException(
                            $"图片网格 {gridHeight} 行 × {gridWidth} 列共 {total} 个单元格，超过单次上限 {CellOperationSupport.MaxCellsPerOperation:0}。");

                    var pixelWidth = ReadNullableDouble(root, "pixel_width") ?? 12.0;
                    var pixelHeight = ReadNullableDouble(root, "pixel_height") ?? 12.0;
                    if (pixelWidth < 1 || pixelWidth > 60)
                        throw new ArgumentException("pixel_width(磅) 必须在 1 到 60 之间。");
                    if (pixelHeight < 2 || pixelHeight > 100)
                        throw new ArgumentException("pixel_height(磅) 必须在 2 到 100 之间。");

                    var hideGridlines = ReadNullableBool(root, "hide_gridlines") ?? true;
                    var palette = ReadPalette(root, "palette");

                    var interpolation = ReadString(root, "interpolation") ?? "nearest";
                    if (!string.Equals(interpolation, "nearest", StringComparison.OrdinalIgnoreCase) &&
                        !string.Equals(interpolation, "bilinear", StringComparison.OrdinalIgnoreCase))
                        throw new ArgumentException("interpolation 只能是 nearest(默认,像素画风格) 或 bilinear(平滑)。");

                    var pixels = SampleToPixels(fullPath, gridWidth, gridHeight, palette, interpolation);
                    var inner = new DrawPixelsOp(ReadString(root, "sheet"), address, pixels, pixelWidth, pixelHeight, hideGridlines);
                    return new DrawFromImageOp(inner, fullPath);
                }
            }

            private static string[,] SampleToPixels(string imagePath, int gridWidth, int gridHeight, List<string> palette, string interpolation)
            {
                var bytes = File.ReadAllBytes(imagePath);
                using (var stream = new MemoryStream(bytes, writable: false))
                using (var source = Image.FromStream(stream))
                using (var scaled = new Bitmap(gridWidth, gridHeight, PixelFormat.Format32bppArgb))
                {
                    using (var graphics = Graphics.FromImage(scaled))
                    {
                        graphics.Clear(Color.Transparent);
                        if (string.Equals(interpolation, "bilinear", StringComparison.OrdinalIgnoreCase))
                        {
                            graphics.InterpolationMode = InterpolationMode.HighQualityBilinear;
                            graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
                            graphics.SmoothingMode = SmoothingMode.HighQuality;
                        }
                        else
                        {
                            graphics.InterpolationMode = InterpolationMode.NearestNeighbor;
                            graphics.PixelOffsetMode = PixelOffsetMode.Half;
                            graphics.SmoothingMode = SmoothingMode.None;
                        }
                        graphics.DrawImage(source, 0, 0, gridWidth, gridHeight);
                    }

                    var paletteColors = palette == null ? null : palette.ConvertAll(ParsePaletteColor);
                    var matrix = new string[gridHeight, gridWidth];
                    for (var r = 0; r < gridHeight; r++)
                    {
                        for (var c = 0; c < gridWidth; c++)
                        {
                            var color = scaled.GetPixel(c, r);
                            if (color.A < 128) { matrix[r, c] = null; continue; }
                            var rgb = paletteColors != null ? NearestPaletteColor(color, paletteColors) : color;
                            matrix[r, c] = ToHex(rgb);
                        }
                    }
                    return matrix;
                }
            }

            private static Color NearestPaletteColor(Color color, List<Color> palette)
            {
                var best = palette[0];
                var bestDistance = int.MaxValue;
                foreach (var candidate in palette)
                {
                    var dr = color.R - candidate.R;
                    var dg = color.G - candidate.G;
                    var db = color.B - candidate.B;
                    var distance = dr * dr + dg * dg + db * db;
                    if (distance < bestDistance)
                    {
                        bestDistance = distance;
                        best = candidate;
                        if (bestDistance == 0) break;
                    }
                }
                return best;
            }

            private static string ToHex(Color color)
            {
                return $"#{color.R:X2}{color.G:X2}{color.B:X2}";
            }

            private static Color ParsePaletteColor(string hex)
            {
                CellOperationSupport.ToOleColor(hex, "palette"); // 校验 #RRGGBB
                return ColorTranslator.FromHtml(hex.Trim());
            }

            private static List<string> ReadPalette(JsonElement root, string name)
            {
                if (!root.TryGetProperty(name, out var element)) return null;
                if (element.ValueKind != JsonValueKind.Array || element.GetArrayLength() == 0)
                    throw new ArgumentException(name + " 必须是非空 #RRGGBB 字符串数组。");
                var list = new List<string>();
                foreach (var item in element.EnumerateArray())
                {
                    if (item.ValueKind != JsonValueKind.String)
                        throw new ArgumentException(name + " 每一项必须是 #RRGGBB 字符串。");
                    var text = item.GetString();
                    if (string.IsNullOrWhiteSpace(text))
                        throw new ArgumentException(name + " 不能包含空字符串。");
                    CellOperationSupport.ToOleColor(text, name);
                    list.Add(text.Trim());
                }
                return list;
            }

            private static int ReadRequiredInt(JsonElement root, string name)
            {
                if (!root.TryGetProperty(name, out var value) || value.ValueKind != JsonValueKind.Number)
                    throw new ArgumentException(name + " 不能为空，必须是整数。");
                var number = value.GetDouble();
                if (number < 1 || number > 100000 || Math.Abs(number - Math.Round(number)) > 0.0001)
                    throw new ArgumentException(name + " 必须是 ≥1 的整数。");
                return (int)number;
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