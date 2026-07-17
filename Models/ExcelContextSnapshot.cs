using System;
using System.Text;
using Microsoft.Office.Interop.Excel;

namespace AgentForExcel.Models
{
    /// <summary>
    /// 当前 Excel 上下文的"快照",作为 LLM 的上下文信息。
    /// 抓取:工作簿名、活动工作表、选中区域、前若干行数据(给模型判断要操作什么)。
    /// 刻意只读少量数据,避免把整表塞进 prompt。
    /// </summary>
    public class ExcelContextSnapshot
    {
        public string WorkbookName { get; set; }
        public string ActiveSheetName { get; set; }
        public string SelectionAddress { get; set; }
        public int RowCount { get; set; }
        public int ColumnCount { get; set; }
        /// <summary>预览数据(前 N 行,二维数组的字符串表示)。</summary>
        public string Preview { get; set; }

        /// <summary>从当前 Excel 实例抓取快照。出错返回最简快照。</summary>
        public static ExcelContextSnapshot Capture(Application excel)
        {
            return Capture(excel, null);
        }

        /// <summary>优先按已锁定选区捕获，避免用户在任务执行期间点击别处导致上下文漂移。</summary>
        public static ExcelContextSnapshot Capture(Application excel, SelectionContext selection)
        {
            var snap = new ExcelContextSnapshot();
            try
            {
                var wb = ResolveWorkbook(excel, selection) ?? excel.ActiveWorkbook;
                snap.WorkbookName = wb?.Name ?? "(无工作簿)";
                Worksheet sheet = null;
                if (wb != null && selection?.IsValid == true)
                {
                    try { sheet = wb.Worksheets[selection.SheetName] as Worksheet; }
                    catch { }
                }
                if (sheet == null) sheet = excel.ActiveSheet as Worksheet;
                snap.ActiveSheetName = sheet?.Name ?? "(无工作表)";

                Range sel = null;
                if (sheet != null && selection?.IsValid == true)
                {
                    try { sel = sheet.get_Range(selection.Address); }
                    catch { }
                }
                if (sel == null) sel = excel.Selection as Range;
                if (sel == null) sel = sheet?.UsedRange;
                if (sel != null)
                {
                    snap.SelectionAddress = sel.Address;
                    snap.RowCount = sel.Rows.Count;
                    snap.ColumnCount = sel.Columns.Count;
                    snap.Preview = ReadPreview(sel, maxRows: 8, maxCols: 8);
                }
            }
            catch (Exception ex)
            {
                snap.Preview = "(读取上下文失败: " + ex.Message + ")";
            }
            return snap;
        }

        private static Workbook ResolveWorkbook(Application excel, SelectionContext selection)
        {
            if (excel == null || selection?.IsValid != true) return null;
            foreach (Workbook workbook in excel.Workbooks)
            {
                try
                {
                    if (!string.IsNullOrWhiteSpace(selection.WorkbookFullName) &&
                        string.Equals(workbook.FullName, selection.WorkbookFullName, StringComparison.OrdinalIgnoreCase))
                        return workbook;
                }
                catch { }
                if (string.Equals(workbook.Name, selection.WorkbookName, StringComparison.OrdinalIgnoreCase))
                    return workbook;
            }
            return null;
        }

        /// <summary>读取区域前 N 行 M 列的文本预览。</summary>
        private static string ReadPreview(Range range, int maxRows, int maxCols)
        {
            int rows = Math.Min(range.Rows.Count, maxRows);
            int cols = Math.Min(range.Columns.Count, maxCols);
            if (rows <= 0 || cols <= 0) return "(空)";

            var small = range.get_Range(range.Cells[1, 1], range.Cells[rows, cols]);
            object[,] values = small.Value2 as object[,];
            if (values == null) return "(无数据)";

            var sb = new StringBuilder();
            for (int r = 1; r <= values.GetLength(0); r++)
            {
                for (int c = 1; c <= values.GetLength(1); c++)
                {
                    var v = values[r, c];
                    sb.Append(v?.ToString() ?? "");
                    if (c < values.GetLength(1)) sb.Append(" | ");
                }
                sb.AppendLine();
            }
            return sb.ToString();
        }

        /// <summary>转成给 LLM 看的简洁文本。</summary>
        public string ToPromptText()
        {
            var sb = new StringBuilder();
            sb.AppendLine("【当前 Excel 上下文】");
            sb.AppendLine($"工作簿: {WorkbookName}");
            sb.AppendLine($"活动工作表: {ActiveSheetName}");
            if (!string.IsNullOrEmpty(SelectionAddress))
                sb.AppendLine($"选中区域: {SelectionAddress}  (约 {RowCount} 行 × {ColumnCount} 列)");
            if (!string.IsNullOrEmpty(Preview))
            {
                sb.AppendLine("数据预览(前若干行):");
                sb.AppendLine(Preview);
            }
            return sb.ToString();
        }
    }
}
