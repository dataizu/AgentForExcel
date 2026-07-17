using System;
using System.Drawing;
using System.Globalization;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;

namespace AgentForExcel.Operations.Cell
{
    /// <summary>单元格写操作共用的目标解析、规模限制和安全校验。</summary>
    internal static class CellOperationSupport
    {
        // 防止一次错误的工具调用锁死 Excel。更大批次应拆成多次、逐次确认。
        internal const double MaxCellsPerOperation = 50000;

        internal static Worksheet GetWorksheet(AppContext context, string sheetName)
        {
            if (context?.Excel?.ActiveWorkbook == null)
                throw new InvalidOperationException("当前没有打开的工作簿。");

            if (string.IsNullOrWhiteSpace(sheetName))
            {
                var activeSheet = context.Excel.ActiveSheet as Worksheet;
                if (activeSheet == null)
                    throw new InvalidOperationException("当前活动对象不是工作表。");
                return activeSheet;
            }

            try
            {
                return (Worksheet)context.Excel.ActiveWorkbook.Worksheets[sheetName.Trim()];
            }
            catch
            {
                throw new ArgumentException("找不到工作表「" + sheetName + "」。");
            }
        }

        internal static Range GetRange(Worksheet sheet, string address)
        {
            if (sheet == null) throw new ArgumentNullException(nameof(sheet));
            if (string.IsNullOrWhiteSpace(address))
                throw new ArgumentException("address 不能为空。");

            Range range;
            try
            {
                range = sheet.get_Range(address.Trim());
            }
            catch
            {
                throw new ArgumentException("无效的单元格区域地址：「" + address + "」。");
            }

            ValidateRangeSize(range);
            return range;
        }

        internal static void ValidateRangeSize(Range range)
        {
            double cellCount;
            try { cellCount = Convert.ToDouble(range.CountLarge, CultureInfo.InvariantCulture); }
            catch { cellCount = Convert.ToDouble(range.Cells.Count, CultureInfo.InvariantCulture); }

            if (cellCount <= 0)
                throw new ArgumentException("目标区域为空。");
            if (cellCount > MaxCellsPerOperation)
                throw new InvalidOperationException(
                    $"单次操作最多允许 {MaxCellsPerOperation:0} 个单元格，当前目标约 {cellCount:0} 个。请拆分区域后重试。");
        }

        internal static int ToOleColor(string value, string fieldName)
        {
            if (string.IsNullOrWhiteSpace(value))
                throw new ArgumentException(fieldName + " 不能为空。");

            var colorText = value.Trim();
            if (!Regex.IsMatch(colorText, "^#[0-9a-fA-F]{6}$"))
                throw new ArgumentException(fieldName + " 必须是 #RRGGBB 格式，例如 #16764A。");

            return ColorTranslator.ToOle(ColorTranslator.FromHtml(colorText));
        }

        internal static void ValidateFormula(string formula)
        {
            if (string.IsNullOrWhiteSpace(formula))
                throw new ArgumentException("formula 不能为空。");
            if (!formula.TrimStart().StartsWith("=", StringComparison.Ordinal))
                throw new ArgumentException("formula 必须以 = 开头。");

            // 阶段 1 只允许工作簿内计算公式，阻止 DDE/外部调用类公式。
            var upper = formula.ToUpperInvariant();
            var blockedTokens = new[]
            {
                "|", "WEBSERVICE(", "RTD(", "CALL(", "REGISTER.ID(", "EXEC(", "SHELL("
            };
            foreach (var token in blockedTokens)
                if (upper.Contains(token))
                    throw new InvalidOperationException("当前安全策略不允许包含外部调用的公式：" + token);
        }
    }
}
