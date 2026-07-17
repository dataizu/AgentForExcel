using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;

namespace AgentForExcel.UI.Controls
{
    /// <summary>
    /// 面向对话回复的轻量 Markdown 查看器。支持标题、段落、粗体、斜体、
    /// 行内代码、列表、代码块和 Markdown 表格，不引入浏览器或第三方渲染器。
    /// </summary>
    public sealed class MarkdownViewer : StackPanel
    {
        public static readonly DependencyProperty TextProperty = DependencyProperty.Register(
            nameof(Text),
            typeof(string),
            typeof(MarkdownViewer),
            new PropertyMetadata(string.Empty, OnTextChanged));

        private static readonly Brush TextBrush = BrushFrom("#23312B");
        private static readonly Brush MutedBrush = BrushFrom("#66736C");
        private static readonly Brush LineBrush = BrushFrom("#E1E7E3");
        private static readonly Brush CodeBrush = BrushFrom("#F3F6F4");

        public MarkdownViewer()
        {
            Orientation = Orientation.Vertical;
        }

        public string Text
        {
            get => (string)GetValue(TextProperty);
            set => SetValue(TextProperty, value);
        }

        private static void OnTextChanged(DependencyObject dependencyObject, DependencyPropertyChangedEventArgs e)
        {
            ((MarkdownViewer)dependencyObject).Render(e.NewValue as string ?? string.Empty);
        }

        private void Render(string markdown)
        {
            Children.Clear();
            var lines = markdown.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
            var index = 0;

            while (index < lines.Length)
            {
                var line = lines[index];
                if (string.IsNullOrWhiteSpace(line))
                {
                    index++;
                    continue;
                }

                if (line.TrimStart().StartsWith("```", StringComparison.Ordinal))
                {
                    var code = new StringBuilder();
                    index++;
                    while (index < lines.Length && !lines[index].TrimStart().StartsWith("```", StringComparison.Ordinal))
                    {
                        if (code.Length > 0) code.AppendLine();
                        code.Append(lines[index]);
                        index++;
                    }
                    if (index < lines.Length) index++;
                    AddCodeBlock(code.ToString());
                    continue;
                }

                if (index + 1 < lines.Length && line.Contains("|") && IsTableSeparator(lines[index + 1]))
                {
                    var tableLines = new List<string> { line };
                    index += 2;
                    while (index < lines.Length && lines[index].Contains("|") && !string.IsNullOrWhiteSpace(lines[index]))
                    {
                        tableLines.Add(lines[index]);
                        index++;
                    }
                    AddTable(tableLines);
                    continue;
                }

                var heading = Regex.Match(line, @"^\s*(#{1,3})\s+(.+)$");
                if (heading.Success)
                {
                    var level = heading.Groups[1].Value.Length;
                    AddRichText(heading.Groups[2].Value, 17 - level, FontWeights.SemiBold, new Thickness(0, 5, 0, 4));
                    index++;
                    continue;
                }

                var bullet = Regex.Match(line, @"^\s*[-*]\s+(.+)$");
                var numbered = Regex.Match(line, @"^\s*(\d+[.)])\s+(.+)$");
                if (bullet.Success || numbered.Success)
                {
                    var marker = bullet.Success ? "•" : numbered.Groups[1].Value;
                    var content = bullet.Success ? bullet.Groups[1].Value : numbered.Groups[2].Value;
                    AddListItem(marker, content);
                    index++;
                    continue;
                }

                var paragraph = new StringBuilder(line.Trim());
                index++;
                while (index < lines.Length && !string.IsNullOrWhiteSpace(lines[index]) && !IsBlockStart(lines, index))
                {
                    paragraph.Append(' ').Append(lines[index].Trim());
                    index++;
                }
                AddRichText(paragraph.ToString(), 12.5, FontWeights.Normal, new Thickness(0, 1, 0, 5));
            }
        }

        private static bool IsBlockStart(string[] lines, int index)
        {
            var line = lines[index];
            if (line.TrimStart().StartsWith("```", StringComparison.Ordinal)) return true;
            if (Regex.IsMatch(line, @"^\s*(#{1,3})\s+")) return true;
            if (Regex.IsMatch(line, @"^\s*[-*]\s+")) return true;
            if (Regex.IsMatch(line, @"^\s*\d+[.)]\s+")) return true;
            return index + 1 < lines.Length && line.Contains("|") && IsTableSeparator(lines[index + 1]);
        }

        private void AddRichText(string text, double fontSize, FontWeight weight, Thickness margin)
        {
            var block = new TextBlock
            {
                FontFamily = new FontFamily("Microsoft YaHei UI"),
                FontSize = fontSize,
                FontWeight = weight,
                Foreground = TextBrush,
                TextWrapping = TextWrapping.Wrap,
                LineHeight = Math.Max(19, fontSize + 7),
                Margin = margin
            };
            AddInlines(block.Inlines, text);
            Children.Add(block);
        }

        private void AddListItem(string marker, string content)
        {
            var grid = new Grid { Margin = new Thickness(0, 1, 0, 4) };
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            grid.Children.Add(new TextBlock
            {
                Text = marker,
                Width = marker == "•" ? 16 : 25,
                FontSize = 12,
                Foreground = BrushFrom("#17734A"),
                Margin = new Thickness(0, 1, 4, 0)
            });
            var body = new TextBlock
            {
                FontFamily = new FontFamily("Microsoft YaHei UI"),
                FontSize = 12.5,
                Foreground = TextBrush,
                TextWrapping = TextWrapping.Wrap,
                LineHeight = 19
            };
            AddInlines(body.Inlines, content);
            Grid.SetColumn(body, 1);
            grid.Children.Add(body);
            Children.Add(grid);
        }

        private void AddCodeBlock(string code)
        {
            var text = new TextBox
            {
                Text = code,
                IsReadOnly = true,
                BorderThickness = new Thickness(0),
                Background = Brushes.Transparent,
                FontFamily = new FontFamily("Consolas"),
                FontSize = 11.5,
                Foreground = TextBrush,
                TextWrapping = TextWrapping.NoWrap,
                Padding = new Thickness(0)
            };
            Children.Add(new Border
            {
                Child = text,
                Background = CodeBrush,
                BorderBrush = LineBrush,
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(8),
                Padding = new Thickness(9, 7, 9, 7),
                Margin = new Thickness(0, 3, 0, 7)
            });
        }

        private void AddTable(IReadOnlyList<string> lines)
        {
            var rows = new List<string[]>();
            foreach (var line in lines) rows.Add(SplitTableRow(line));
            var columnCount = 0;
            foreach (var row in rows) columnCount = Math.Max(columnCount, row.Length);
            if (columnCount == 0) return;

            var grid = new Grid();
            for (var column = 0; column < columnCount; column++)
                grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(96) });
            for (var row = 0; row < rows.Count; row++)
            {
                grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
                for (var column = 0; column < columnCount; column++)
                {
                    var value = column < rows[row].Length ? rows[row][column] : string.Empty;
                    var cell = new Border
                    {
                        BorderBrush = LineBrush,
                        BorderThickness = new Thickness(column == 0 ? 1 : 0, row == 0 ? 1 : 0, 1, 1),
                        Background = row == 0 ? BrushFrom("#EDF5F0") : (row % 2 == 0 ? BrushFrom("#FAFCFA") : Brushes.White),
                        Padding = new Thickness(7, 5, 7, 5),
                        Child = new TextBlock
                        {
                            Text = value,
                            FontSize = 11.5,
                            FontWeight = row == 0 ? FontWeights.SemiBold : FontWeights.Normal,
                            Foreground = TextBrush,
                            TextTrimming = TextTrimming.CharacterEllipsis,
                            ToolTip = value
                        }
                    };
                    Grid.SetRow(cell, row);
                    Grid.SetColumn(cell, column);
                    grid.Children.Add(cell);
                }
            }

            Children.Add(new ScrollViewer
            {
                Content = grid,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                VerticalScrollBarVisibility = ScrollBarVisibility.Disabled,
                Margin = new Thickness(0, 3, 0, 7)
            });
        }

        private static bool IsTableSeparator(string line)
        {
            var cells = SplitTableRow(line);
            if (cells.Length == 0) return false;
            foreach (var cell in cells)
                if (!Regex.IsMatch(cell.Trim(), @"^:?-{3,}:?$")) return false;
            return true;
        }

        private static string[] SplitTableRow(string line)
        {
            var trimmed = line.Trim().Trim('|');
            if (string.IsNullOrWhiteSpace(trimmed)) return new string[0];
            var cells = trimmed.Split('|');
            for (var i = 0; i < cells.Length; i++) cells[i] = cells[i].Trim();
            return cells;
        }

        private static void AddInlines(InlineCollection inlines, string text)
        {
            var index = 0;
            while (index < text.Length)
            {
                var bold = text.IndexOf("**", index, StringComparison.Ordinal);
                var code = text.IndexOf('`', index);
                var italic = text.IndexOf('*', index);
                var next = MinPositive(bold, code, italic);
                if (next < 0)
                {
                    inlines.Add(new Run(text.Substring(index)));
                    break;
                }
                if (next > index) inlines.Add(new Run(text.Substring(index, next - index)));

                if (next == bold)
                {
                    var end = text.IndexOf("**", bold + 2, StringComparison.Ordinal);
                    if (end < 0) { inlines.Add(new Run(text.Substring(next))); break; }
                    inlines.Add(new Run(text.Substring(bold + 2, end - bold - 2)) { FontWeight = FontWeights.SemiBold });
                    index = end + 2;
                }
                else if (next == code)
                {
                    var end = text.IndexOf('`', code + 1);
                    if (end < 0) { inlines.Add(new Run(text.Substring(next))); break; }
                    inlines.Add(new Run(text.Substring(code + 1, end - code - 1))
                    {
                        FontFamily = new FontFamily("Consolas"),
                        Background = CodeBrush,
                        Foreground = BrushFrom("#145F3D")
                    });
                    index = end + 1;
                }
                else
                {
                    var end = text.IndexOf('*', italic + 1);
                    if (end < 0) { inlines.Add(new Run(text.Substring(next))); break; }
                    inlines.Add(new Run(text.Substring(italic + 1, end - italic - 1)) { FontStyle = FontStyles.Italic });
                    index = end + 1;
                }
            }
        }

        private static int MinPositive(params int[] values)
        {
            var result = -1;
            foreach (var value in values)
                if (value >= 0 && (result < 0 || value < result)) result = value;
            return result;
        }

        private static Brush BrushFrom(string color)
            => (Brush)new BrushConverter().ConvertFrom(color);
    }
}
