using System;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;

namespace AgentForExcel.Operations.Analysis
{
    internal static class AnalysisSheetSupport
    {
        internal static Worksheet CreateUniqueWorksheet(AppContext context, string preferredName)
        {
            if (context?.Excel?.ActiveWorkbook == null)
                throw new InvalidOperationException("当前没有打开的工作簿。");

            var workbook = context.Excel.ActiveWorkbook;
            var baseName = SanitizeWorksheetName(
                string.IsNullOrWhiteSpace(preferredName) ? "Agent分析" : preferredName.Trim());
            var candidate = baseName;
            var suffix = 2;
            while (WorksheetExists(workbook, candidate))
            {
                var suffixText = suffix++.ToString();
                candidate = baseName.Substring(0, Math.Min(baseName.Length, 31 - suffixText.Length)) + suffixText;
            }

            var sheet = (Worksheet)workbook.Worksheets.Add(After: workbook.Worksheets[workbook.Worksheets.Count]);
            sheet.Name = candidate;
            sheet.Tab.Color = Cell.CellOperationSupport.ToOleColor("#168653", "tab_color");
            return sheet;
        }

        internal static void DeleteWorksheetSilently(AppContext context, Worksheet worksheet)
        {
            if (worksheet == null) return;
            var application = context?.Excel;
            var previousAlerts = application?.DisplayAlerts ?? true;
            try
            {
                if (application != null) application.DisplayAlerts = false;
                worksheet.Delete();
            }
            catch { }
            finally
            {
                if (application != null) application.DisplayAlerts = previousAlerts;
            }
        }

        private static bool WorksheetExists(Workbook workbook, string name)
        {
            foreach (Worksheet sheet in workbook.Worksheets)
                if (string.Equals(sheet.Name, name, StringComparison.OrdinalIgnoreCase)) return true;
            return false;
        }

        private static string SanitizeWorksheetName(string value)
        {
            var cleaned = Regex.Replace(value, @"[\\/:?*\[\]]", "_").Trim(' ', '\'');
            if (string.IsNullOrWhiteSpace(cleaned)) cleaned = "Agent分析";
            return cleaned.Length <= 31 ? cleaned : cleaned.Substring(0, 31);
        }
    }
}
