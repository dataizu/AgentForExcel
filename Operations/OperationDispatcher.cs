using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading.Tasks;
using AgentForExcel.Models;
using AgentForExcel.Services;

namespace AgentForExcel.Operations
{
    /// <summary>
    /// 操作派发器。维护 工具名→工厂 的映射,把 LLM 返回的 JSON 指令解析为
    /// IOperation 列表,经 SafetyGuard 确认后逐个执行。
    /// </summary>
    public class OperationDispatcher
    {
        private readonly Dictionary<string, IOperationFactory> _factories = new Dictionary<string, IOperationFactory>();

        /// <summary>是否已注册该工具。</summary>
        public bool IsRegistered(string toolName) => _factories.ContainsKey(toolName);

        /// <summary>注册一个操作工厂(启动时统一调用)。</summary>
        public void Register(IOperationFactory factory)
        {
            if (factory == null) throw new ArgumentNullException(nameof(factory));
            _factories[factory.ToolName] = factory;
        }

        /// <summary>
        /// 解析并执行 LLM 下发的一批操作。
        /// </summary>
        /// <param name="operationPayloads">LLM 返回的各 tool_call arguments(JSON 字符串)。
        ///     注意:此处简化为只传 arguments;完整 tool_call 还含工具名。
        ///     为对接实际协议,见 ParseWithToolName 重载。</param>
        /// <param name="confirmCallback">写操作的确认回调,返回 true 才执行。</param>
        /// <returns>每条操作的执行结果文本(展示给用户)。</returns>
        public Task<IReadOnlyList<string>> ExecuteAsync(
            IReadOnlyList<OperationCall> calls,
            Func<string, bool> confirmCallback)
        {
            var results = new List<string>();
            foreach (var call in calls)
            {
                var toolName = call?.ToolName ?? "unknown";
                var operationTimer = Stopwatch.StartNew();
                var outcome = "completed";
                try
                {
                    // Excel COM 对象必须留在创建它的 UI/STA 线程上执行。
                    // 不使用 Task.Run，避免跨线程 COM 调用导致随机失败或 Excel 卡死。
                    string result = ExecuteOne(call, confirmCallback);
                    outcome = ClassifyOutcome(result);
                    results.Add(result);
                }
                catch (Exception ex)
                {
                    outcome = "exception";
                    results.Add($"[{toolName}] 执行失败：{ex.Message}");
                }
                finally
                {
                    operationTimer.Stop();
                    PerformanceLogger.Log(
                        "operation",
                        operationTimer.ElapsedMilliseconds,
                        "tool=" + toolName + "|outcome=" + outcome);
                }
            }
            return Task.FromResult<IReadOnlyList<string>>(results);
        }

        private string ExecuteOne(OperationCall call, Func<string, bool> confirmCallback)
        {
            var editionReason = GetEditionDisabledReason(call?.ToolName);
            if (editionReason != null)
                return $"[{call.ToolName}] 已阻止：{editionReason}";

            var disabledReason = GetDisabledReason(call?.ToolName, ThisAddIn.App?.Settings);
            if (disabledReason != null)
                return $"[{call.ToolName}] 已阻止：{disabledReason}。可在“设置 → 执行与安全”中重新开启。";

            // 1) 查工厂
            if (!_factories.TryGetValue(call.ToolName, out var factory))
                return $"未知工具：{call.ToolName}（该能力尚未实现）";

            // 2) 解析 JSON 为 IOperation
            IOperation op;
            try { op = factory.Parse(call.ArgumentsJson); }
            catch (Exception ex) { return $"[{call.ToolName}] 参数解析失败：{ex.Message}"; }

            // 3) 写操作按权限策略判断：安全范围自动执行，高风险操作仍需确认。
            var decision = ThisAddIn.App?.Permissions?.Evaluate(call, op);
            var requiresConfirmation = op.IsWriteOperation &&
                                       (decision?.RequiresConfirmation ?? ThisAddIn.App.Settings.RequireConfirmOnWrite);
            if (requiresConfirmation)
            {
                var reason = string.IsNullOrWhiteSpace(decision?.Reason) ? "写操作需要确认" : decision.Reason;
                bool approved = confirmCallback?.Invoke(op.Describe() + "\n\n需要确认的原因：" + reason) ?? false;
                if (!approved)
                    return $"已跳过（未确认）：{op.Describe()}";
            }
            else if (op.IsWriteOperation)
            {
                ThisAddIn.Log("安全自动执行：" + call.ToolName + "；" + (decision?.Reason ?? "策略允许"));
            }

            // 4) 执行
            var context = ThisAddIn.App;
            string outcome = op.Execute(context);
            return string.IsNullOrEmpty(outcome) ? op.Describe() : outcome;
        }

        private static string GetDisabledReason(string toolName, Models.UserSettings settings)
        {
            if (string.IsNullOrWhiteSpace(toolName) || settings == null) return null;
            if (toolName.StartsWith("pq_", StringComparison.OrdinalIgnoreCase) && !settings.EnablePowerQuery)
                return "Power Query 能力已关闭";
            if (toolName.StartsWith("pp_", StringComparison.OrdinalIgnoreCase) && !settings.EnablePowerPivot)
                return "Power Pivot / DAX 能力已关闭";
            if (toolName.StartsWith("vba_", StringComparison.OrdinalIgnoreCase) && !settings.EnableVba)
                return "受控 VBA 能力已关闭";
            return null;
        }

        private static string GetEditionDisabledReason(string toolName)
        {
            return EditionPolicy.IsToolAvailable(toolName, ProductEditionInfo.Current)
                ? null
                : EditionPolicy.GetUnavailableReason(toolName, ProductEditionInfo.Current);
        }

        private static string ClassifyOutcome(string result)
        {
            if (string.IsNullOrWhiteSpace(result)) return "completed";
            if (result.StartsWith("已跳过", StringComparison.Ordinal)) return "skipped";
            if (result.StartsWith("未知工具", StringComparison.Ordinal)) return "unknown_tool";
            if (result.Contains("已阻止")) return "blocked";
            if (result.Contains("参数解析失败")) return "parse_failed";
            return "completed";
        }
    }

    /// <summary>一次工具调用的结构化表示(工具名 + 参数 JSON)。</summary>
    public sealed class OperationCall
    {
        /// <summary>模型返回的 tool_call id；回传工具结果时必须原样关联。</summary>
        public string CallId { get; set; }
        public string ToolName { get; set; }
        public string ArgumentsJson { get; set; }
    }
}
