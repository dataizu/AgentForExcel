using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text.Json;

namespace AgentForExcel.Services
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public sealed class AgentAutomationService
    {
        public string HealthCheck()
        {
            var context = ThisAddIn.App;
            if (context == null) throw new InvalidOperationException("Agent for Excel 尚未完成初始化。");
            return new Operations.SystemCheck.AgentSelfCheckOp().Execute(context);
        }

        public string CapabilityCheck()
        {
            var context = ThisAddIn.App;
            if (context?.Dispatcher == null) throw new InvalidOperationException("Agent 工具派发器尚未完成初始化。");
            var required = new[]
            {
                "task_plan", "task_step_update",
                "cell_read_range", "cell_write_range", "cell_fill_formula", "cell_format_range",
                "data_profile", "analysis_create_view", "chart_create", "pivot_create", "dashboard_create",
                "pq_list_queries", "pq_create_from_range", "pq_load_to_sheet", "pq_refresh",
                "pp_list_model", "pp_add_query_to_model", "pp_refresh_model", "pp_add_relationship",
                "pp_add_measure", "pp_create_model_pivot",
                "vba_preview_safe", "vba_execute_safe", "agent_self_check"
            };
            var missing = new List<string>();
            foreach (var tool in required)
                if (!context.Dispatcher.IsRegistered(tool)) missing.Add(tool);
            return JsonSerializer.Serialize(new
            {
                required_count = required.Length,
                registered_count = required.Length - missing.Count,
                all_registered = missing.Count == 0,
                missing
            });
        }

        public string RunReadOnlyTool(string toolName, string argumentsJson)
        {
            var allowed = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "agent_self_check", "cell_read_range", "data_profile", "pq_list_queries", "pp_list_model", "vba_preview_safe"
            };
            if (!allowed.Contains(toolName ?? string.Empty))
                throw new InvalidOperationException("COM 诊断接口只允许执行只读工具。");
            var context = ThisAddIn.App;
            if (context?.Dispatcher == null) throw new InvalidOperationException("Agent 工具派发器尚未完成初始化。");
            var calls = new[]
            {
                new Operations.OperationCall
                {
                    CallId = "automation-diagnostic",
                    ToolName = toolName,
                    ArgumentsJson = string.IsNullOrWhiteSpace(argumentsJson) ? "{}" : argumentsJson
                }
            };
            var results = context.Dispatcher.ExecuteAsync(calls, _ => false).GetAwaiter().GetResult();
            return results.Count == 0 ? string.Empty : results[0];
        }

        public string Version() => typeof(AgentAutomationService).Assembly.GetName().Version.ToString();
    }
}
