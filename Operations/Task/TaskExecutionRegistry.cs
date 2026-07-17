using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;

namespace AgentForExcel.Operations.Tasking
{
    public sealed class TaskStepSnapshot
    {
        public int Index { get; set; }
        public string Title { get; set; }
        public string Status { get; set; }
        public string Detail { get; set; }
    }

    public sealed class TaskPlanSnapshot
    {
        public string Title { get; set; }
        public string UserRequest { get; set; }
        public List<TaskStepSnapshot> Steps { get; set; } = new List<TaskStepSnapshot>();
        public List<string> SuccessCriteria { get; set; } = new List<string>();
        public bool IsComplete { get; set; }
        public int CompletedSteps { get; set; }
        public int TotalSteps { get; set; }
    }

    /// <summary>当前 Agent 任务的轻量状态机。一次对话执行开始时重置。</summary>
    public static class TaskExecutionRegistry
    {
        public const string PayloadPrefix = "__AGENT_TASK_PLAN__";
        private static readonly object SyncRoot = new object();
        private static TaskPlanSnapshot _current;
        private static string _userRequest;

        public static void BeginRun(string userRequest)
        {
            lock (SyncRoot)
            {
                _userRequest = userRequest ?? string.Empty;
                _current = null;
            }
        }

        public static TaskPlanSnapshot SetPlan(string title, IList<string> steps, IList<string> successCriteria)
        {
            if (steps == null || steps.Count == 0)
                throw new ArgumentException("任务计划至少需要一个步骤。", nameof(steps));
            if (steps.Count > 12)
                throw new ArgumentException("任务计划最多支持 12 个步骤。", nameof(steps));

            lock (SyncRoot)
            {
                _current = new TaskPlanSnapshot
                {
                    Title = string.IsNullOrWhiteSpace(title) ? "执行计划" : title.Trim(),
                    UserRequest = _userRequest,
                    Steps = steps.Select((step, index) => new TaskStepSnapshot
                    {
                        Index = index + 1,
                        Title = step?.Trim(),
                        Status = index == 0 ? "in_progress" : "pending"
                    }).ToList(),
                    SuccessCriteria = successCriteria?.Where(x => !string.IsNullOrWhiteSpace(x))
                        .Select(x => x.Trim()).Take(8).ToList() ?? new List<string>()
                };
                Recalculate(_current);
                return Clone(_current);
            }
        }

        public static TaskPlanSnapshot UpdateStep(int stepIndex, string status, string detail)
        {
            lock (SyncRoot)
            {
                if (_current == null) throw new InvalidOperationException("尚未创建任务计划。请先调用 task_plan。");
                if (stepIndex < 1 || stepIndex > _current.Steps.Count)
                    throw new ArgumentException("step_index 超出任务计划范围。");
                var normalized = NormalizeStatus(status);
                if ((normalized == "completed" || normalized == "failed") && string.IsNullOrWhiteSpace(detail))
                    throw new ArgumentException("标记 completed 或 failed 时必须在 detail 中提供核验证据或失败原因。");
                var step = _current.Steps[stepIndex - 1];
                step.Status = normalized;
                step.Detail = string.IsNullOrWhiteSpace(detail) ? null : detail.Trim();

                if (normalized == "completed" && stepIndex < _current.Steps.Count)
                {
                    var next = _current.Steps[stepIndex];
                    if (next.Status == "pending") next.Status = "in_progress";
                }
                Recalculate(_current);
                return Clone(_current);
            }
        }

        public static TaskPlanSnapshot GetSnapshot()
        {
            lock (SyncRoot) return Clone(_current);
        }

        public static bool HasActivePlan
        {
            get { lock (SyncRoot) return _current != null; }
        }

        public static bool IsComplete
        {
            get { lock (SyncRoot) return _current != null && _current.IsComplete; }
        }

        public static string DescribeIncompleteSteps()
        {
            lock (SyncRoot)
            {
                if (_current == null) return string.Empty;
                var remaining = _current.Steps
                    .Where(step => step.Status != "completed")
                    .Select(step => step.Index + ". " + step.Title + "（" + step.Status + "）");
                return string.Join("；", remaining);
            }
        }

        public static string Serialize(TaskPlanSnapshot snapshot)
        {
            return PayloadPrefix + JsonSerializer.Serialize(snapshot ?? GetSnapshot());
        }

        private static string NormalizeStatus(string status)
        {
            switch ((status ?? string.Empty).Trim().ToLowerInvariant())
            {
                case "pending":
                case "in_progress":
                case "completed":
                case "failed":
                    return status.Trim().ToLowerInvariant();
                default:
                    throw new ArgumentException("status 仅支持 pending、in_progress、completed 或 failed。");
            }
        }

        private static void Recalculate(TaskPlanSnapshot plan)
        {
            plan.TotalSteps = plan.Steps.Count;
            plan.CompletedSteps = plan.Steps.Count(step => step.Status == "completed");
            plan.IsComplete = plan.TotalSteps > 0 && plan.CompletedSteps == plan.TotalSteps;
        }

        private static TaskPlanSnapshot Clone(TaskPlanSnapshot source)
        {
            if (source == null) return null;
            return new TaskPlanSnapshot
            {
                Title = source.Title,
                UserRequest = source.UserRequest,
                SuccessCriteria = new List<string>(source.SuccessCriteria ?? new List<string>()),
                IsComplete = source.IsComplete,
                CompletedSteps = source.CompletedSteps,
                TotalSteps = source.TotalSteps,
                Steps = source.Steps.Select(step => new TaskStepSnapshot
                {
                    Index = step.Index,
                    Title = step.Title,
                    Status = step.Status,
                    Detail = step.Detail
                }).ToList()
            };
        }
    }
}
