using System;
using System.Collections.Generic;
using System.Text.Json;
using AgentForExcel.Models;
using AgentForExcel.Services;
using Microsoft.Office.Interop.Excel;

namespace AgentForExcel.Operations.SystemCheck
{
    public sealed class AgentSelfCheckOp : IOperation
    {
        internal const string Prefix = "__AGENT_SELF_CHECK__";
        public string ToolName => "agent_self_check";
        public bool IsWriteOperation => false;
        public string Describe() => "检查 Agent for Excel 的当前运行环境";

        public string Execute(AppContext context)
        {
            var application = context?.Excel;
            if (application == null) throw new InvalidOperationException("Excel 应用程序上下文不可用。");
            var warnings = new List<string>();
            var runtime = RuntimeEnvironmentInspector.Capture();
            var edition = ProductEditionInfo.Current;
            var apiConfigured = !string.IsNullOrWhiteSpace(context?.Settings?.ApiKey);
            var workbook = application.ActiveWorkbook;
            if (workbook == null)
            {
                warnings.Add("当前没有打开的工作簿；数据操作能力暂不可用。");
                return Prefix + JsonSerializer.Serialize(new
                {
                    excel_version = application.Version,
                    workbook_open = false,
                    power_query = false,
                    power_pivot = false,
                    vba_project_access = false,
                    edition = ProductEditionInfo.Id,
                    edition_name = ProductEditionInfo.CurrentDisplayName,
                    edition_description = ProductEditionInfo.Description(edition),
                    api_configured = apiConfigured,
                    runtime,
                    warnings
                });
            }

            var saved = !string.IsNullOrWhiteSpace(workbook.Path);
            if (!saved) warnings.Add("工作簿尚未保存；受控 VBA 不会执行，因为无法创建可靠的执行前备份。");

            var powerQuery = false;
            var queryCount = 0;
            try
            {
                dynamic queries = ((dynamic)workbook).Queries;
                queryCount = Convert.ToInt32(queries.Count);
                powerQuery = true;
            }
            catch { warnings.Add("当前 Excel 版本未暴露 Power Query 工作簿对象模型。"); }

            var powerPivot = false;
            var modelTableCount = 0;
            try
            {
                dynamic model = ((dynamic)workbook).Model;
                dynamic tables = model.ModelTables;
                modelTableCount = Convert.ToInt32(tables.Count);
                powerPivot = true;
            }
            catch { warnings.Add("当前工作簿或 Excel 版本未启用 Power Pivot 数据模型对象模型。"); }

            var vbaAccess = false;
            try
            {
                dynamic project = ((dynamic)workbook).VBProject;
                vbaAccess = project != null;
            }
            catch { }
            if (!vbaAccess)
                warnings.Add("未开启“信任对 VBA 工程对象模型的访问”；受控 VBA 可预览，但执行会被安全拦截。");

            return Prefix + JsonSerializer.Serialize(new
            {
                excel_version = application.Version,
                workbook_open = true,
                workbook = workbook.Name,
                workbook_saved = saved,
                file_format = Convert.ToInt32(workbook.FileFormat),
                power_query = powerQuery,
                query_count = queryCount,
                power_pivot = powerPivot,
                model_table_count = modelTableCount,
                vba_project_access = vbaAccess,
                edition = ProductEditionInfo.Id,
                edition_name = ProductEditionInfo.CurrentDisplayName,
                edition_description = ProductEditionInfo.Description(edition),
                api_configured = apiConfigured,
                runtime,
                warnings
            });
        }

        public sealed class Factory : IOperationFactory
        {
            public string ToolName => "agent_self_check";
            public IOperation Parse(string argumentsJson) => new AgentSelfCheckOp();
        }
    }
}
