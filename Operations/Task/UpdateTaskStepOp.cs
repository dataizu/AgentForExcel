using System;
using System.Text.Json;

namespace AgentForExcel.Operations.Tasking
{
    public sealed class UpdateTaskStepOp : IOperation
    {
        private readonly int _stepIndex;
        private readonly string _status;
        private readonly string _detail;

        private UpdateTaskStepOp(int stepIndex, string status, string detail)
        {
            _stepIndex = stepIndex;
            _status = status;
            _detail = detail;
        }

        public string ToolName => "task_step_update";
        public bool IsWriteOperation => false;
        public string Describe() => "更新任务步骤 " + _stepIndex + " 为 " + _status;

        public string Execute(AppContext context)
        {
            return TaskExecutionRegistry.Serialize(
                TaskExecutionRegistry.UpdateStep(_stepIndex, _status, _detail));
        }

        public sealed class Factory : IOperationFactory
        {
            public string ToolName => "task_step_update";

            public IOperation Parse(string argumentsJson)
            {
                using (var document = JsonDocument.Parse(argumentsJson))
                {
                    var root = document.RootElement;
                    if (!root.TryGetProperty("step_index", out var indexValue) ||
                        indexValue.ValueKind != JsonValueKind.Number || !indexValue.TryGetInt32(out var stepIndex))
                        throw new ArgumentException("step_index 必须是整数。");
                    if (!root.TryGetProperty("status", out var statusValue) || statusValue.ValueKind != JsonValueKind.String)
                        throw new ArgumentException("status 不能为空。");
                    var detail = root.TryGetProperty("detail", out var detailValue) && detailValue.ValueKind == JsonValueKind.String
                        ? detailValue.GetString()
                        : null;
                    return new UpdateTaskStepOp(stepIndex, statusValue.GetString(), detail);
                }
            }
        }
    }
}
