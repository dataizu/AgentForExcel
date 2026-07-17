using System;
using System.Collections.Generic;
using System.Text.Json;

namespace AgentForExcel.Operations.Tasking
{
    public sealed class PlanTaskOp : IOperation
    {
        private readonly string _title;
        private readonly string[] _steps;
        private readonly string[] _successCriteria;

        private PlanTaskOp(string title, string[] steps, string[] successCriteria)
        {
            _title = title;
            _steps = steps;
            _successCriteria = successCriteria;
        }

        public string ToolName => "task_plan";
        public bool IsWriteOperation => false;
        public string Describe() => "创建任务执行计划「" + (_title ?? "执行计划") + "」";

        public string Execute(AppContext context)
        {
            return TaskExecutionRegistry.Serialize(
                TaskExecutionRegistry.SetPlan(_title, _steps, _successCriteria));
        }

        public sealed class Factory : IOperationFactory
        {
            public string ToolName => "task_plan";

            public IOperation Parse(string argumentsJson)
            {
                using (var document = JsonDocument.Parse(argumentsJson))
                {
                    var root = document.RootElement;
                    var steps = ReadArray(root, "steps", true);
                    var criteria = ReadArray(root, "success_criteria", false);
                    var title = root.TryGetProperty("title", out var titleValue) && titleValue.ValueKind == JsonValueKind.String
                        ? titleValue.GetString()?.Trim()
                        : null;
                    return new PlanTaskOp(title, steps, criteria);
                }
            }

            private static string[] ReadArray(JsonElement root, string name, bool required)
            {
                if (!root.TryGetProperty(name, out var value))
                {
                    if (required) throw new ArgumentException(name + " 不能为空。");
                    return new string[0];
                }
                if (value.ValueKind != JsonValueKind.Array)
                    throw new ArgumentException(name + " 必须是字符串数组。");
                var result = new List<string>();
                foreach (var item in value.EnumerateArray())
                {
                    if (item.ValueKind != JsonValueKind.String || string.IsNullOrWhiteSpace(item.GetString()))
                        throw new ArgumentException(name + " 不能包含空值。");
                    result.Add(item.GetString().Trim());
                }
                if (required && result.Count == 0) throw new ArgumentException(name + " 至少需要一项。");
                return result.ToArray();
            }
        }
    }
}
