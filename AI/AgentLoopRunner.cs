using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using AgentForExcel.Operations;
using AgentForExcel.Operations.Tasking;

namespace AgentForExcel.AI
{
    /// <summary>LLM 单轮回复。</summary>
    public sealed class LlmReply
    {
        public string Text { get; set; }

        public IReadOnlyList<OperationCall> Operations { get; set; } = new List<OperationCall>();

        public bool HasOperations => Operations != null && Operations.Count > 0;
    }

    /// <summary>
    /// 一条协议级对话消息。除普通 user/assistant 文本外，也能保存 assistant 的
    /// tool_calls 和逐条 tool 执行结果，确保模型可以在下一轮继续完成任务。
    /// </summary>
    public sealed class ChatTurn
    {
        public string Role { get; set; }
        public string Content { get; set; }
        public IReadOnlyList<OperationCall> ToolCalls { get; set; }
        public string ToolCallId { get; set; }
    }

    public sealed class AgentRunResult
    {
        public bool Completed { get; set; }
        public int Rounds { get; set; }
        public int ToolCallCount { get; set; }
        public int CompletionCheckCount { get; set; }
        public string StopReason { get; set; }
    }

    /// <summary>
    /// 多轮 Agent 执行器：模型提出工具调用后，执行工具、把结果按协议回传，
    /// 再自动请求模型继续，直到模型返回不含工具调用的最终答复。
    /// </summary>
    public static class AgentLoopRunner
    {
        public const int MaxRounds = 12;
        public const int MaxToolCalls = 36;
        public const int MaxCompletionNudges = 3;

        public static async Task<AgentRunResult> RunAsync(
            string userMessage,
            IList<ChatTurn> history,
            Func<IReadOnlyList<ChatTurn>, Task<LlmReply>> requestAsync,
            Func<IReadOnlyList<OperationCall>, Task<IReadOnlyList<string>>> executeAsync,
            Action<int> onRoundStarting,
            Action<LlmReply, int> onAssistantReply,
            Action<OperationCall, string, int, int> onToolResult)
        {
            if (string.IsNullOrWhiteSpace(userMessage))
                throw new ArgumentException("用户消息不能为空。", nameof(userMessage));
            if (history == null) throw new ArgumentNullException(nameof(history));
            if (requestAsync == null) throw new ArgumentNullException(nameof(requestAsync));
            if (executeAsync == null) throw new ArgumentNullException(nameof(executeAsync));

            history.Add(new ChatTurn { Role = "user", Content = userMessage });
            if (!(IsContinuationMessage(userMessage) && TaskExecutionRegistry.HasActivePlan && !TaskExecutionRegistry.IsComplete))
                TaskExecutionRegistry.BeginRun(userMessage);
            var toolCallCount = 0;
            var emptyReplyCount = 0;
            var completionCheckCount = 0;
            var lastToolFailureSummary = string.Empty;
            var pendingCellWriteVerification = false;

            for (var round = 1; round <= MaxRounds; round++)
            {
                onRoundStarting?.Invoke(round);
                var reply = await requestAsync(new List<ChatTurn>(history));
                if (reply == null)
                    throw new InvalidOperationException("大模型返回了空响应。");

                EnsureToolCallIds(reply.Operations);
                history.Add(new ChatTurn
                {
                    Role = "assistant",
                    Content = reply.Text ?? string.Empty,
                    ToolCalls = reply.Operations
                });
                onAssistantReply?.Invoke(reply, round);

                if (!reply.HasOperations)
                {
                    if (string.IsNullOrWhiteSpace(reply.Text) && emptyReplyCount < 1 && round < MaxRounds)
                    {
                        emptyReplyCount++;
                        history.Add(new ChatTurn
                        {
                            Role = "user",
                            Content = "请基于已有上下文和工具结果，直接输出面向用户的最终答复；不要只输出思考过程。"
                        });
                        continue;
                    }

                    var continuationReason = GetContinuationReason(
                        reply.Text, lastToolFailureSummary, pendingCellWriteVerification);
                    if (!string.IsNullOrWhiteSpace(continuationReason))
                    {
                        if (completionCheckCount < MaxCompletionNudges && round < MaxRounds)
                        {
                            completionCheckCount++;
                            history.Add(new ChatTurn
                            {
                                Role = "user",
                                Content = "完成检查未通过：" + continuationReason +
                                          "。请立即基于现有上下文继续调用必要工具、核验结果并完成剩余工作；不要只解释计划，也不要重复已经成功的操作。"
                            });
                            continue;
                        }
                        return new AgentRunResult
                        {
                            Completed = false,
                            Rounds = round,
                            ToolCallCount = toolCallCount,
                            CompletionCheckCount = completionCheckCount,
                            StopReason = TaskExecutionRegistry.HasActivePlan && !TaskExecutionRegistry.IsComplete
                                ? "incomplete_plan"
                                : "completion_check_limit"
                        };
                    }
                    return new AgentRunResult
                    {
                        Completed = true,
                        Rounds = round,
                        ToolCallCount = toolCallCount,
                        CompletionCheckCount = completionCheckCount,
                        StopReason = "completed"
                    };
                }

                if (toolCallCount + reply.Operations.Count > MaxToolCalls)
                {
                    return new AgentRunResult
                    {
                        Completed = false,
                        Rounds = round,
                        ToolCallCount = toolCallCount,
                        CompletionCheckCount = completionCheckCount,
                        StopReason = "tool_call_limit"
                    };
                }

                var results = await executeAsync(reply.Operations) ?? new List<string>();
                toolCallCount += reply.Operations.Count;
                var failures = new List<string>();

                for (var i = 0; i < reply.Operations.Count; i++)
                {
                    var result = i < results.Count
                        ? results[i]
                        : $"[{reply.Operations[i].ToolName}] 未返回执行结果。";
                    history.Add(new ChatTurn
                    {
                        Role = "tool",
                        ToolCallId = reply.Operations[i].CallId,
                        Content = result ?? string.Empty
                    });
                    var failed = IsFailureResult(result);
                    if (failed) failures.Add(result);
                    else if (RequiresCellVerification(reply.Operations[i].ToolName)) pendingCellWriteVerification = true;
                    else if (reply.Operations[i].ToolName == "cell_read_range") pendingCellWriteVerification = false;
                    onToolResult?.Invoke(reply.Operations[i], result, i, reply.Operations.Count);
                }
                lastToolFailureSummary = failures.Count == 0
                    ? string.Empty
                    : string.Join("；", failures);
            }

            return new AgentRunResult
            {
                Completed = false,
                Rounds = MaxRounds,
                ToolCallCount = toolCallCount,
                CompletionCheckCount = completionCheckCount,
                StopReason = "round_limit"
            };
        }

        private static string GetContinuationReason(
            string replyText, string lastToolFailureSummary, bool pendingCellWriteVerification)
        {
            if (LooksLikeBlockedReply(replyText)) return null;
            if (TaskExecutionRegistry.HasActivePlan && !TaskExecutionRegistry.IsComplete)
                return "任务计划仍有未完成步骤：" + TaskExecutionRegistry.DescribeIncompleteSteps();
            if (!string.IsNullOrWhiteSpace(lastToolFailureSummary))
                return "上一轮仍有未处理的工具失败：" + lastToolFailureSummary;
            if (pendingCellWriteVerification)
                return "本次写值或公式操作尚未通过 cell_read_range 回读核验";
            if (LooksLikeInterimReply(replyText))
                return "当前回复仍停留在准备或计划阶段，没有交付用户要求的结果";
            return null;
        }

        private static bool LooksLikeInterimReply(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return false;
            var value = text.Trim();
            var markers = new[]
            {
                "我需要先", "我先", "让我先", "接下来我会", "接下来将", "现在需要", "正在读取",
                "正在查看", "正在处理", "准备读取", "准备分析", "稍等", "请稍候"
            };
            foreach (var marker in markers)
                if (value.IndexOf(marker, StringComparison.OrdinalIgnoreCase) >= 0) return true;
            return false;
        }

        private static bool LooksLikeBlockedReply(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return false;
            var markers = new[] { "需要你提供", "请提供", "需要用户确认", "未获得授权", "无法继续", "当前没有打开" };
            foreach (var marker in markers)
                if (text.IndexOf(marker, StringComparison.OrdinalIgnoreCase) >= 0) return true;
            return false;
        }

        private static bool IsFailureResult(string result)
        {
            if (string.IsNullOrWhiteSpace(result)) return true;
            var markers = new[] { "执行失败", "参数解析失败", "未知工具", "已跳过", "未确认", "未返回执行结果" };
            foreach (var marker in markers)
                if (result.IndexOf(marker, StringComparison.OrdinalIgnoreCase) >= 0) return true;
            return false;
        }

        private static bool IsContinuationMessage(string message)
        {
            var value = (message ?? string.Empty).Trim();
            return value == "继续" || value == "继续执行" || value == "接着做" || value == "请继续" ||
                   value.StartsWith("继续完成", StringComparison.OrdinalIgnoreCase);
        }

        private static bool RequiresCellVerification(string toolName)
        {
            return toolName == "cell_write_range" || toolName == "cell_fill_formula";
        }

        private static void EnsureToolCallIds(IReadOnlyList<OperationCall> calls)
        {
            if (calls == null) return;
            foreach (var call in calls)
            {
                if (string.IsNullOrWhiteSpace(call.CallId))
                    call.CallId = "call_" + Guid.NewGuid().ToString("N");
            }
        }
    }
}
