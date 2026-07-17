using System;
using System.Text.Json;
using System.Text.RegularExpressions;
using AgentForExcel.Models;
using AgentForExcel.Operations;
using Excel = Microsoft.Office.Interop.Excel;

namespace AgentForExcel.Services
{
    public sealed class PermissionDecision
    {
        private PermissionDecision(bool requiresConfirmation, string reason)
        {
            RequiresConfirmation = requiresConfirmation;
            Reason = reason;
        }

        public bool RequiresConfirmation { get; }
        public string Reason { get; }

        public static PermissionDecision Allow(string reason) => new PermissionDecision(false, reason);
        public static PermissionDecision Confirm(string reason) => new PermissionDecision(true, reason);
    }

    /// <summary>
    /// 将“是否确认”从简单开关升级为范围感知的权限策略。
    /// 只在任务锁定选区或明确的新工作表输出范围内自动执行，危险能力始终询问。
    /// </summary>
    public sealed class PermissionPolicyService
    {
        private static readonly Regex ExternalWorkbookFormula = new Regex(
            @"\[[^\]]+\.(xlsx|xlsm|xlsb|xls)\]", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private readonly AppContext _context;

        public PermissionPolicyService(AppContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
        }

        public PermissionDecision Evaluate(OperationCall call, IOperation operation)
        {
            if (operation == null || !operation.IsWriteOperation)
                return PermissionDecision.Allow("只读操作");

            var settings = _context.Settings;
            if (settings == null || string.Equals(settings.AutomationMode, "ask_every_time", StringComparison.OrdinalIgnoreCase))
                return PermissionDecision.Confirm("当前设置为每次询问");

            var tool = call?.ToolName ?? operation.ToolName ?? string.Empty;
            if (IsAlwaysConfirmTool(tool))
                return PermissionDecision.Confirm(AlwaysConfirmReason(tool));

            JsonDocument document = null;
            try
            {
                document = JsonDocument.Parse(string.IsNullOrWhiteSpace(call?.ArgumentsJson) ? "{}" : call.ArgumentsJson);
                var root = document.RootElement;

                if (tool == "pq_create_from_range")
                    return ReadBool(root, "replace_existing")
                        ? PermissionDecision.Confirm("将覆盖现有 Power Query 查询定义")
                        : PermissionDecision.Confirm("创建 Power Query 可能新增源表对象和工作簿查询定义");

                if (settings.AutoAllowNewSheetOutputs && IsSafeNewSheetOutput(tool, root))
                    return PermissionDecision.Allow("结果输出到新的工作表，源数据保持不变");

                if (tool == "cell_format_range")
                {
                    if (!settings.AutoAllowFormattingInSelection)
                        return PermissionDecision.Confirm("未开启选区内自动格式化");
                    if (ReadBool(root, "autofit_columns") || ReadBool(root, "autofit_rows"))
                        return PermissionDecision.Confirm("自动列宽或行高会影响整个行列");
                    return EvaluateLockedSelectionWrite(root, false, settings.AutoWriteMaxCells);
                }

                if (tool == "cell_write_range" || tool == "cell_fill_formula")
                {
                    if (!settings.AutoAllowSelectedBlankWrites)
                        return PermissionDecision.Confirm("未开启锁定选区内自动写入");
                    if (tool == "cell_fill_formula")
                    {
                        var formula = ReadString(root, "formula");
                        if (!string.IsNullOrWhiteSpace(formula) && ExternalWorkbookFormula.IsMatch(formula))
                            return PermissionDecision.Confirm("公式包含外部工作簿引用");
                    }
                    return EvaluateLockedSelectionWrite(root, true, settings.AutoWriteMaxCells);
                }

                return PermissionDecision.Confirm("该写操作不在安全自动化白名单内");
            }
            catch (Exception ex)
            {
                return PermissionDecision.Confirm("无法完成安全范围校验：" + ex.Message);
            }
            finally
            {
                document?.Dispose();
            }
        }

        private PermissionDecision EvaluateLockedSelectionWrite(JsonElement root, bool requireBlank, int maxCells)
        {
            var selection = _context.Selection?.Locked;
            if (selection == null || !selection.IsValid)
                return PermissionDecision.Confirm("没有锁定任务选区");
            if (selection.IsMultiArea)
                return PermissionDecision.Confirm("多区域选区不允许自动写入");

            var activeWorkbook = _context.Excel.ActiveWorkbook;
            if (activeWorkbook == null || !WorkbookMatches(activeWorkbook, selection))
                return PermissionDecision.Confirm("当前活动工作簿与锁定选区不一致");

            var requestedSheet = ReadString(root, "sheet");
            if (!string.IsNullOrWhiteSpace(requestedSheet) &&
                !string.Equals(requestedSheet, selection.SheetName, StringComparison.OrdinalIgnoreCase))
                return PermissionDecision.Confirm("目标工作表不在锁定选区内");

            var address = ReadString(root, "address");
            if (string.IsNullOrWhiteSpace(address))
                return PermissionDecision.Confirm("目标区域不明确");

            var sheet = activeWorkbook.Worksheets[selection.SheetName] as Excel.Worksheet;
            var lockedRange = sheet?.get_Range(selection.Address);
            var target = sheet?.get_Range(address);
            if (sheet == null || lockedRange == null || target == null)
                return PermissionDecision.Confirm("无法解析锁定选区或目标区域");

            if (root.TryGetProperty("values", out var values) &&
                target.Rows.Count == 1 && target.Columns.Count == 1)
            {
                GetValueShape(values, out var rows, out var columns);
                if (rows > 1 || columns > 1) target = target.get_Resize(rows, columns);
            }

            var targetCells = Convert.ToInt64(target.CountLarge);
            if (targetCells > Math.Max(1, maxCells))
                return PermissionDecision.Confirm($"目标共 {targetCells:#,##0} 个单元格，超过自动写入上限 {maxCells:#,##0}");

            var intersection = _context.Excel.Intersect(lockedRange, target);
            if (intersection == null || Convert.ToInt64(intersection.CountLarge) != targetCells)
                return PermissionDecision.Confirm("目标区域超出锁定选区");

            if (requireBlank && Convert.ToDouble(_context.Excel.WorksheetFunction.CountA(target)) > 0)
                return PermissionDecision.Confirm("目标区域包含已有值或公式");

            return PermissionDecision.Allow(requireBlank
                ? "目标位于锁定选区内且当前为空白"
                : "格式化范围完全位于锁定选区内");
        }

        private bool IsSafeNewSheetOutput(string tool, JsonElement root)
        {
            switch (tool)
            {
                case "analysis_create_view":
                case "dashboard_create":
                case "pq_load_to_sheet":
                case "pp_create_model_pivot":
                    return true;
                case "chart_create":
                case "pivot_create":
                    var destination = ReadString(root, "destination_sheet");
                    return string.IsNullOrWhiteSpace(destination) || !WorksheetExists(destination);
                default:
                    return false;
            }
        }

        private bool WorksheetExists(string name)
        {
            var workbook = _context.Excel.ActiveWorkbook;
            if (workbook == null || string.IsNullOrWhiteSpace(name)) return false;
            try
            {
                var ignored = workbook.Worksheets[name];
                return ignored != null;
            }
            catch
            {
                return false;
            }
        }

        private static bool IsAlwaysConfirmTool(string tool)
        {
            return tool == "vba_execute_safe" || tool == "pp_add_query_to_model" ||
                   tool == "pp_add_relationship" || tool == "pp_add_measure" ||
                   tool == "pp_refresh_model" || tool == "pq_refresh";
        }

        private static string AlwaysConfirmReason(string tool)
        {
            if (tool.StartsWith("vba_", StringComparison.OrdinalIgnoreCase)) return "VBA 会执行工作簿级自动化";
            if (tool == "pq_refresh" || tool == "pp_refresh_model") return "刷新会改写已加载的查询或模型结果";
            return "该操作会修改 Power Pivot 数据模型、关系或 DAX 定义";
        }

        private static bool WorkbookMatches(Excel.Workbook workbook, SelectionContext selection)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(selection.WorkbookFullName) &&
                    string.Equals(workbook.FullName, selection.WorkbookFullName, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            catch { }
            return string.Equals(workbook.Name, selection.WorkbookName, StringComparison.OrdinalIgnoreCase);
        }

        private static string ReadString(JsonElement root, string name)
        {
            return root.TryGetProperty(name, out var value) && value.ValueKind == JsonValueKind.String
                ? value.GetString()?.Trim()
                : null;
        }

        private static bool ReadBool(JsonElement root, string name)
        {
            return root.TryGetProperty(name, out var value) &&
                   (value.ValueKind == JsonValueKind.True || value.ValueKind == JsonValueKind.False) &&
                   value.GetBoolean();
        }

        private static void GetValueShape(JsonElement values, out int rows, out int columns)
        {
            rows = 1;
            columns = 1;
            if (values.ValueKind != JsonValueKind.Array) return;
            var outer = values.GetArrayLength();
            if (outer == 0) { rows = 0; columns = 0; return; }
            var first = values[0];
            if (first.ValueKind == JsonValueKind.Array)
            {
                rows = outer;
                columns = first.GetArrayLength();
            }
            else
            {
                columns = outer;
            }
        }
    }
}
