using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text.Json;
using Microsoft.Office.Interop.Excel;

namespace AgentForExcel.Operations.Macro
{
    internal sealed class SafeVbaPreview
    {
        public string Token { get; set; }
        public string WorkbookKey { get; set; }
        public string Recipe { get; set; }
        public string Summary { get; set; }
        public string Code { get; set; }
        public string OutputPath { get; set; }
        public DateTime ExpiresAtUtc { get; set; }
    }

    internal static class SafeVbaSupport
    {
        internal const string PreviewPrefix = "__AGENT_VBA_PREVIEW__";
        private static readonly ConcurrentDictionary<string, SafeVbaPreview> Previews =
            new ConcurrentDictionary<string, SafeVbaPreview>(StringComparer.OrdinalIgnoreCase);

        internal static Workbook GetWorkbook(AppContext context)
        {
            var workbook = context?.Excel?.ActiveWorkbook;
            if (workbook == null) throw new InvalidOperationException("当前没有打开的工作簿。");
            return workbook;
        }

        internal static string WorkbookKey(Workbook workbook)
        {
            var fullName = string.Empty;
            try { fullName = workbook.FullName; } catch { }
            return string.IsNullOrWhiteSpace(fullName) ? workbook.Name : fullName;
        }

        internal static SafeVbaPreview CreatePreview(Workbook workbook, string recipe, string outputPath)
        {
            recipe = (recipe ?? string.Empty).Trim().ToLowerInvariant();
            string summary;
            string body;
            switch (recipe)
            {
                case "refresh_all":
                    summary = "刷新当前工作簿的全部连接、Power Query、数据模型和透视表";
                    body = "    ThisWorkbook.RefreshAll\r\n" +
                           "    Application.CalculateUntilAsyncQueriesDone\r\n";
                    break;
                case "autofit_used_ranges":
                    summary = "自动调整所有可见工作表的已用区域列宽，并把最大列宽限制为 40";
                    body = "    Dim ws As Worksheet\r\n" +
                           "    Dim col As Range\r\n" +
                           "    For Each ws In ThisWorkbook.Worksheets\r\n" +
                           "        If ws.Visible = xlSheetVisible Then\r\n" +
                           "            ws.UsedRange.Columns.AutoFit\r\n" +
                           "            For Each col In ws.UsedRange.Columns\r\n" +
                           "                If col.ColumnWidth > 40 Then col.ColumnWidth = 40\r\n" +
                           "            Next col\r\n" +
                           "        End If\r\n" +
                           "    Next ws\r\n";
                    break;
                case "export_active_sheet_pdf":
                    outputPath = NormalizePdfPath(workbook, outputPath);
                    summary = "把当前活动工作表导出为 PDF：" + outputPath;
                    body = "    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=\"" +
                           EscapeVba(outputPath) + "\", Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False\r\n";
                    break;
                default:
                    throw new ArgumentException("recipe 仅支持 refresh_all、autofit_used_ranges 或 export_active_sheet_pdf。");
            }

            var token = Guid.NewGuid().ToString("N");
            var macroName = "AgentSafe_" + token.Substring(0, 12);
            var code = "Option Explicit\r\n\r\nPublic Sub " + macroName + "()\r\n" + body + "End Sub\r\n";
            var preview = new SafeVbaPreview
            {
                Token = token,
                WorkbookKey = WorkbookKey(workbook),
                Recipe = recipe,
                Summary = summary,
                Code = code,
                OutputPath = outputPath,
                ExpiresAtUtc = DateTime.UtcNow.AddMinutes(20)
            };
            Previews[token] = preview;
            CleanupExpired();
            return preview;
        }

        internal static SafeVbaPreview ConsumePreview(Workbook workbook, string token)
        {
            if (string.IsNullOrWhiteSpace(token) || !Previews.TryRemove(token.Trim(), out var preview))
                throw new ArgumentException("预览令牌无效或已经使用，请重新预览宏。");
            if (preview.ExpiresAtUtc < DateTime.UtcNow)
                throw new InvalidOperationException("宏预览已经过期，请重新生成预览。");
            if (!string.Equals(preview.WorkbookKey, WorkbookKey(workbook), StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("当前工作簿与宏预览时不一致，已拒绝执行。");
            return preview;
        }

        internal static string Execute(Workbook workbook, SafeVbaPreview preview)
        {
            if (string.IsNullOrWhiteSpace(workbook.Path))
                throw new InvalidOperationException("为保证可回滚，请先保存当前工作簿，再执行受控 VBA。");

            var backupDirectory = Path.Combine(workbook.Path, "AgentBackups");
            Directory.CreateDirectory(backupDirectory);
            var extension = Path.GetExtension(workbook.Name);
            var safeName = Path.GetFileNameWithoutExtension(workbook.Name);
            var backupPath = Path.Combine(backupDirectory,
                safeName + "_before_vba_" + DateTime.Now.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture) + extension);
            workbook.SaveCopyAs(backupPath);

            dynamic project;
            try
            {
                project = ((dynamic)workbook).VBProject;
                if (project == null) throw new InvalidOperationException("VBProject unavailable");
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException(
                    "Excel 未允许访问 VBA 工程。请在“文件 → 选项 → 信任中心 → 宏设置”中启用“信任对 VBA 工程对象模型的访问”，然后重试。", ex);
            }

            dynamic component = null;
            var moduleName = "AgentSafe" + preview.Token.Substring(0, 8);
            var macroName = "AgentSafe_" + preview.Token.Substring(0, 12);
            try
            {
                component = project.VBComponents.Add(1);
                component.Name = moduleName;
                component.CodeModule.AddFromString(preview.Code);
                workbook.Application.Run("'" + workbook.Name.Replace("'", "''") + "'!" + moduleName + "." + macroName);
                AppendAudit(workbook, preview, "成功", backupPath);
                return "已执行受控 VBA「" + preview.Recipe + "」。执行前备份：" + backupPath + "。";
            }
            catch (Exception ex)
            {
                AppendAudit(workbook, preview, "失败：" + ex.Message, backupPath);
                throw new InvalidOperationException("受控 VBA 执行失败；原工作簿未自动覆盖，执行前备份位于：" + backupPath + "。", ex);
            }
            finally
            {
                if (component != null)
                {
                    try { project.VBComponents.Remove(component); } catch { }
                }
            }
        }

        private static string NormalizePdfPath(Workbook workbook, string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                var directory = string.IsNullOrWhiteSpace(workbook.Path) ? Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) : workbook.Path;
                dynamic activeSheet = workbook.Application.ActiveSheet;
                path = Path.Combine(directory, workbook.Name + "_" + Convert.ToString(activeSheet.Name) + ".pdf");
            }
            path = Path.GetFullPath(path.Trim());
            if (!string.Equals(Path.GetExtension(path), ".pdf", StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("PDF 导出路径必须以 .pdf 结尾。");
            return path;
        }

        private static string EscapeVba(string value) => (value ?? string.Empty).Replace("\"", "\"\"");

        private static void AppendAudit(Workbook workbook, SafeVbaPreview preview, string status, string backupPath)
        {
            try
            {
                Worksheet sheet = null;
                foreach (Worksheet candidate in workbook.Worksheets)
                {
                    if (string.Equals(candidate.Name, "__AgentAudit", StringComparison.OrdinalIgnoreCase))
                    {
                        sheet = candidate;
                        break;
                    }
                }
                if (sheet == null)
                {
                    sheet = (Worksheet)workbook.Worksheets.Add(After: workbook.Worksheets.Item[workbook.Worksheets.Count]);
                    sheet.Name = "__AgentAudit";
                    sheet.Range["A1:F1"].Value2 = new object[,] { { "时间", "类型", "操作", "状态", "令牌", "备份" } };
                    sheet.Visible = XlSheetVisibility.xlSheetVeryHidden;
                }
                dynamic cells = sheet.Cells;
                dynamic lastCell = cells[sheet.Rows.Count, 1];
                var row = Convert.ToInt32(lastCell.End[XlDirection.xlUp].Row) + 1;
                sheet.Range[sheet.Cells[row, 1], sheet.Cells[row, 6]].Value2 = new object[,]
                {
                    { DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), "VBA", preview.Recipe, status, preview.Token, backupPath }
                };
            }
            catch { }
        }

        private static void CleanupExpired()
        {
            foreach (var item in Previews)
                if (item.Value.ExpiresAtUtc < DateTime.UtcNow) Previews.TryRemove(item.Key, out _);
        }
    }

    public sealed class PreviewSafeVbaOp : IOperation
    {
        private readonly string _recipe;
        private readonly string _outputPath;
        private PreviewSafeVbaOp(string recipe, string outputPath) { _recipe = recipe; _outputPath = outputPath; }
        public string ToolName => "vba_preview_safe";
        public bool IsWriteOperation => false;
        public string Describe() => "预览受控 VBA 配方「" + _recipe + "」";
        public string Execute(AppContext context)
        {
            var preview = SafeVbaSupport.CreatePreview(SafeVbaSupport.GetWorkbook(context), _recipe, _outputPath);
            return SafeVbaSupport.PreviewPrefix + JsonSerializer.Serialize(new
            {
                preview_token = preview.Token,
                recipe = preview.Recipe,
                summary = preview.Summary,
                expires_at = preview.ExpiresAtUtc.ToString("O"),
                safeguards = new[] { "仅执行内置白名单配方", "执行前创建工作簿副本", "执行后移除临时模块", "写入隐藏审计日志" },
                code = preview.Code
            });
        }
        public sealed class Factory : IOperationFactory
        {
            public string ToolName => "vba_preview_safe";
            public IOperation Parse(string argumentsJson)
            {
                using (var document = JsonDocument.Parse(argumentsJson))
                {
                    var root = document.RootElement;
                    if (!root.TryGetProperty("recipe", out var recipe) || recipe.ValueKind != JsonValueKind.String || string.IsNullOrWhiteSpace(recipe.GetString()))
                        throw new ArgumentException("recipe 不能为空。");
                    var output = root.TryGetProperty("output_path", out var value) && value.ValueKind == JsonValueKind.String ? value.GetString() : null;
                    return new PreviewSafeVbaOp(recipe.GetString(), output);
                }
            }
        }
    }

    public sealed class ExecuteSafeVbaOp : IOperation
    {
        private readonly string _token;
        private ExecuteSafeVbaOp(string token) { _token = token; }
        public string ToolName => "vba_execute_safe";
        public bool IsWriteOperation => true;
        public string Describe() => "执行已预览并授权的受控 VBA";
        public string Execute(AppContext context)
        {
            var workbook = SafeVbaSupport.GetWorkbook(context);
            return SafeVbaSupport.Execute(workbook, SafeVbaSupport.ConsumePreview(workbook, _token));
        }
        public sealed class Factory : IOperationFactory
        {
            public string ToolName => "vba_execute_safe";
            public IOperation Parse(string argumentsJson)
            {
                using (var document = JsonDocument.Parse(argumentsJson))
                {
                    if (!document.RootElement.TryGetProperty("preview_token", out var value) || value.ValueKind != JsonValueKind.String || string.IsNullOrWhiteSpace(value.GetString()))
                        throw new ArgumentException("preview_token 不能为空。");
                    return new ExecuteSafeVbaOp(value.GetString().Trim());
                }
            }
        }
    }
}
