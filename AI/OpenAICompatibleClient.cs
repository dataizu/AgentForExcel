using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using AgentForExcel.Models;
using AgentForExcel.Operations;
using System.Text.Json;

namespace AgentForExcel.AI
{
    /// <summary>
    /// OpenAI 兼容协议客户端(/v1/chat/completions)。
    /// 智谱 GLM、DeepSeek、通义千问均提供此协议,换后端只改 BaseUrl/ApiKey/Model。
    ///
    /// 用 HttpClient 直接发 POST,不依赖任何 SDK —— 体积小、行为可控、
    /// 也方便在 VSTO 的 .NET Framework 4.8 环境中部署。
    /// </summary>
    public class OpenAICompatibleClient : ILLMClient, IDisposable
    {
        private static readonly HttpClient Http = new HttpClient { Timeout = TimeSpan.FromSeconds(180) };
        private UserSettings _settings;

        static OpenAICompatibleClient()
        {
            // VSTO 运行在 .NET Framework 中，显式保证现代模型接口需要的 TLS 1.2。
            ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
        }

        public OpenAICompatibleClient(UserSettings settings)
        {
            _settings = settings ?? throw new ArgumentNullException(nameof(settings));
        }

        /// <summary>运行时刷新连接参数(用户改了设置后调用)。</summary>
        public void ReloadConfig(UserSettings settings) => _settings = settings;

        public Task<LlmReply> ChatAsync(string userMessage, ExcelContextSnapshot excelContext)
            => ChatAsync(userMessage, excelContext, history: null);

        public async Task<LlmReply> ChatAsync(string userMessage, ExcelContextSnapshot excelContext, IReadOnlyList<ChatTurn> history)
        {
            if (string.IsNullOrWhiteSpace(_settings.ApiKey))
                throw new InvalidOperationException("未配置 API Key,请先在设置中填写。");

            var messages = BuildMessages(userMessage, excelContext, history);

            // 2) 组装请求体
            var body = new
            {
                model = _settings.Model,
                messages,
                temperature = _settings.Temperature,
                // function calling 工具:阶段 2 起注册实际工具,当前为空数组
                tools = ToolDefinitions.Tools
            };

            string jsonBody = JsonSerializer.Serialize(body, JsonOpts);
            string url = _settings.BaseUrl.TrimEnd('/') + "/chat/completions";

            // 3) 发请求
            var req = new HttpRequestMessage(HttpMethod.Post, url);
            req.Headers.Add("Authorization", "Bearer " + _settings.ApiKey);
            req.Content = new StringContent(jsonBody, Encoding.UTF8, "application/json");

            using (var resp = await Http.SendAsync(req))
            {
                string respText = await resp.Content.ReadAsStringAsync();
                if (!resp.IsSuccessStatusCode)
                    throw new InvalidOperationException(
                        $"大模型请求失败 ({(int)resp.StatusCode} {resp.StatusCode}):\n{TruncateForDisplay(respText)}");
                return ParseReply(respText);
            }
        }

        public async Task<LlmReply> ChatStreamingAsync(
            string userMessage,
            ExcelContextSnapshot excelContext,
            IReadOnlyList<ChatTurn> history,
            System.Threading.CancellationToken cancellationToken,
            Action<string> onTextDelta)
        {
            if (string.IsNullOrWhiteSpace(_settings.ApiKey))
                throw new InvalidOperationException("未配置 API Key,请先在设置中填写。");

            var body = new Dictionary<string, object>
            {
                ["model"] = _settings.Model,
                ["messages"] = BuildMessages(userMessage, excelContext, history),
                ["temperature"] = _settings.Temperature,
                ["tools"] = ToolDefinitions.Tools,
                ["stream"] = true
            };

            // 智谱 GLM-4.7/4.6 需要显式开启流式工具调用。
            if (IsZhipuEndpoint())
            {
                body["tool_stream"] = true;
                body["thinking"] = new { type = "enabled" };
            }

            var url = _settings.BaseUrl.TrimEnd('/') + "/chat/completions";
            var jsonBody = JsonSerializer.Serialize(body, JsonOpts);
            Exception lastConnectionError = null;
            TransientHttpException lastTransient = null;
            const int maxAttempts = 3;
            for (var attempt = 0; attempt < maxAttempts; attempt++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                try
                {
                    return await SendStreamingOnceAsync(
                        url,
                        jsonBody,
                        cancellationToken,
                        onTextDelta,
                        forceNewConnection: attempt > 0);
                }
                catch (TaskCanceledException ex)
                {
                    // TaskCanceledException 派生自 OperationCanceledException,必须先判:
                    // 是用户取消就按取消向上抛,否则视为 HttpClient 整体超时。
                    cancellationToken.ThrowIfCancellationRequested();
                    throw new InvalidOperationException("模型请求超过 180 秒，已停止等待。请缩小分析范围后重试。", ex);
                }
                catch (OperationCanceledException)
                {
                    throw; // 用户主动停止/流空闲超时不参与重试,直接向上传播
                }
                catch (TransientHttpException ex)
                {
                    // 限流/网关错误:优先按 Retry-After 等待,指数退避,最多 maxAttempts 次。
                    lastTransient = ex;
                    if (attempt == maxAttempts - 1) break;
                    var delaySeconds = Math.Min(ex.RetryAfterSeconds ?? (attempt + 1) * 2, 30);
                    await Task.Delay(TimeSpan.FromSeconds(delaySeconds), cancellationToken);
                    continue;
                }
                catch (HttpRequestException ex)
                {
                    lastConnectionError = ex;
                }
                catch (IOException ex)
                {
                    throw new InvalidOperationException(
                        "模型流式响应中途断开：" + GetDeepestMessage(ex) + "。请直接重试本条消息。",
                        ex);
                }

                if (attempt < maxAttempts - 1)
                    await Task.Delay(400, cancellationToken);
            }

            if (lastTransient != null)
                throw new InvalidOperationException(
                    "模型服务限流或暂时不可用（" + lastTransient.Status + "），已自动重试 " + maxAttempts +
                    " 次仍失败。请稍等片刻再重试本条消息。", lastTransient);

            var host = new Uri(url).Host;
            throw new InvalidOperationException(
                $"无法连接模型服务 {host}，自动重试仍失败：{GetDeepestMessage(lastConnectionError)}",
                lastConnectionError);
        }

        /// <summary>可重试的服务端临时故障(限流/网关),携带服务端建议的等待秒数。</summary>
        private sealed class TransientHttpException : Exception
        {
            public int Status { get; }
            public int? RetryAfterSeconds { get; }

            public TransientHttpException(int status, int? retryAfterSeconds)
                : base("模型服务临时故障 (" + status + ")")
            {
                Status = status;
                RetryAfterSeconds = retryAfterSeconds;
            }
        }

        /// <summary>错误响应体截断,避免长 JSON 撑爆聊天气泡。</summary>
        private static string TruncateForDisplay(string text, int maxLength = 500)
        {
            if (string.IsNullOrEmpty(text) || text.Length <= maxLength) return text;
            return text.Substring(0, maxLength) + "…(已截断)";
        }

        private async Task<LlmReply> SendStreamingOnceAsync(
            string url,
            string jsonBody,
            System.Threading.CancellationToken cancellationToken,
            Action<string> onTextDelta,
            bool forceNewConnection)
        {
            using (var request = new HttpRequestMessage(HttpMethod.Post, url))
            {
                request.Headers.Add("Authorization", "Bearer " + _settings.ApiKey);
                request.Headers.Add("Accept", "text/event-stream");
                if (forceNewConnection) request.Headers.ConnectionClose = true;
                request.Content = new StringContent(jsonBody, Encoding.UTF8, "application/json");

                using (var response = await Http.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, cancellationToken))
                {
                    if (!response.IsSuccessStatusCode)
                    {
                        var error = await response.Content.ReadAsStringAsync();
                        var status = (int)response.StatusCode;
                        if (status == 429 || status == 502 || status == 503 || status == 504)
                        {
                            int? retryAfter = null;
                            try
                            {
                                if (response.Headers.TryGetValues("Retry-After", out var values))
                                    foreach (var value in values)
                                    {
                                        int seconds;
                                        if (int.TryParse(value, out seconds))
                                        {
                                            retryAfter = seconds;
                                            break;
                                        }
                                    }
                            }
                            catch { }
                            throw new TransientHttpException(status, retryAfter);
                        }
                        throw new InvalidOperationException(
                            $"大模型请求失败 ({status} {response.StatusCode}):\n{TruncateForDisplay(error)}");
                    }

                    var accumulator = new StreamingReplyAccumulator();
                    var receivedChunk = false;
                    using (var stream = await response.Content.ReadAsStreamAsync())
                    using (var reader = new StreamReader(stream, Encoding.UTF8))
                    {
                        // .NET Framework 的 ReadLineAsync 没有取消重载:
                        // 用"空闲超时 + 取消时主动关闭流"的组合 —— 关流会让
                        // 挂起的 ReadLineAsync 立即抛出,从而中止整个读取循环。
                        // 关流注册必须挂在 linked(用户取消 ∪ 空闲超时)上:
                        // 只挂 idleCts 的话用户取消无法中断挂起的读取。
                        using (var idleCts = new System.Threading.CancellationTokenSource())
                        using (var linked = System.Threading.CancellationTokenSource.CreateLinkedTokenSource(
                            cancellationToken, idleCts.Token))
                        using (linked.Token.Register(() => { try { stream.Dispose(); } catch { } }))
                        {
                            var idleTimer = new System.Threading.Timer(
                                _ => { try { idleCts.Cancel(); } catch { } },
                                null, StreamIdleTimeoutMilliseconds, System.Threading.Timeout.Infinite);
                            try
                            {
                                string line;
                                while (true)
                                {
                                    // 正在持续输出数据时,关流不会发生;循环内显式响应取消,
                                    // 让"停止"在长回复流式阶段立即生效。
                                    cancellationToken.ThrowIfCancellationRequested();
                                    line = await reader.ReadLineAsync();
                                    // 每收到一行就重置空闲计时;连接中断(返回 null)时立即取消计时器,
                                    // 避免 finally 阶段误触发"空闲超时"。
                                    idleTimer.Change(line == null
                                        ? System.Threading.Timeout.Infinite
                                        : StreamIdleTimeoutMilliseconds,
                                        System.Threading.Timeout.Infinite);
                                    if (line == null) break;

                                    if (string.IsNullOrWhiteSpace(line)) continue;

                                    var payload = line.StartsWith("data:", StringComparison.OrdinalIgnoreCase)
                                        ? line.Substring(5).TrimStart()
                                        : line.Trim();
                                    if (payload == "[DONE]") break;
                                    if (!payload.StartsWith("{", StringComparison.Ordinal)) continue;

                                    receivedChunk = true;
                                    var delta = accumulator.Consume(payload);
                                    if (!string.IsNullOrEmpty(delta)) onTextDelta?.Invoke(delta);
                                }

                                if (idleCts.IsCancellationRequested && !cancellationToken.IsCancellationRequested)
                                    throw CreateIdleTimeoutException();
                            }
                            catch (ObjectDisposedException)
                            {
                                if (cancellationToken.IsCancellationRequested)
                                    throw new OperationCanceledException("已停止生成。", cancellationToken);
                                if (idleCts.IsCancellationRequested)
                                    throw CreateIdleTimeoutException();
                                throw;
                            }
                            finally
                            {
                                idleTimer.Dispose();
                            }
                        }
                    }

                    if (!receivedChunk)
                        throw new IOException("模型连接已建立，但没有收到可识别的流式数据。");
                    var reply = accumulator.BuildReply();
                    ThisAddIn.Log($"LLM 流式完成: textLength={reply.Text?.Length ?? 0}, toolCalls={reply.Operations.Count}");
                    return reply;
                }
            }
        }

        /// <summary>流式读取阶段两行数据之间允许的最长等待;超时视为服务端停滞。</summary>
        private const int StreamIdleTimeoutMilliseconds = 180000;

        /// <summary>
        /// 空闲超时用 InvalidOperationException 而非 OCE:OCE 会被 UI 统一显示成
        /// "已停止生成",把服务端停滞误报成用户主动停止;专用异常走通用错误分支,
        /// 用户能看到真实原因。
        /// </summary>
        private static InvalidOperationException CreateIdleTimeoutException()
        {
            return new InvalidOperationException(
                "模型流式响应超过 " + StreamIdleTimeoutMilliseconds / 1000 +
                " 秒没有新数据，已自动中断。请直接重试本条消息。");
        }

        private static string GetDeepestMessage(Exception exception)
        {
            if (exception == null) return "未知连接错误";
            while (exception.InnerException != null) exception = exception.InnerException;
            return string.IsNullOrWhiteSpace(exception.Message) ? exception.GetType().Name : exception.Message;
        }

        private bool IsZhipuEndpoint()
        {
            return (_settings.BaseUrl ?? string.Empty).IndexOf("bigmodel.cn", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   (_settings.ProviderName ?? string.Empty).IndexOf("智谱", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        /// <summary>解析 OpenAI 兼容协议的响应,提取文本与 function calling 的参数。</summary>
        private LlmReply ParseReply(string respText)
        {
            using (var doc = JsonDocument.Parse(respText))
            {
                var root = doc.RootElement;
                // 取 choices[0].message
                var message = root.GetProperty("choices")[0].GetProperty("message");

                string text = message.TryGetProperty("content", out var contentEl) && contentEl.ValueKind == JsonValueKind.String
                    ? contentEl.GetString()
                    : "";

                var ops = new List<OperationCall>();
                // 若模型返回了 tool_calls(function calling),逐个提取 工具名+参数JSON
                if (message.TryGetProperty("tool_calls", out var toolCalls) && toolCalls.ValueKind == JsonValueKind.Array)
                {
                    foreach (var tc in toolCalls.EnumerateArray())
                    {
                        if (tc.TryGetProperty("function", out var fn) &&
                            fn.TryGetProperty("name", out var nameEl) &&
                            nameEl.ValueKind == JsonValueKind.String &&
                            fn.TryGetProperty("arguments", out var args) &&
                            args.ValueKind == JsonValueKind.String)
                        {
                            ops.Add(new OperationCall
                            {
                                CallId = tc.TryGetProperty("id", out var idEl) && idEl.ValueKind == JsonValueKind.String
                                    ? idEl.GetString()
                                    : null,
                                ToolName = nameEl.GetString(),
                                ArgumentsJson = args.GetString()
                            });
                        }
                    }
                }

                return new LlmReply { Text = text ?? "", Operations = ops };
            }
        }

        private static object ToProtocolMessage(ChatTurn turn)
        {
            if (turn == null)
                return new { role = "assistant", content = string.Empty };

            if (string.Equals(turn.Role, "assistant", StringComparison.OrdinalIgnoreCase) &&
                turn.ToolCalls != null && turn.ToolCalls.Count > 0)
            {
                var toolCalls = new List<object>();
                foreach (var call in turn.ToolCalls)
                {
                    toolCalls.Add(new
                    {
                        id = call.CallId,
                        type = "function",
                        function = new
                        {
                            name = call.ToolName,
                            arguments = call.ArgumentsJson ?? "{}"
                        }
                    });
                }

                return new
                {
                    role = "assistant",
                    content = turn.Content ?? string.Empty,
                    tool_calls = toolCalls
                };
            }

            if (string.Equals(turn.Role, "tool", StringComparison.OrdinalIgnoreCase))
            {
                var content = turn.Content ?? string.Empty;
                if (content.Length > MaxToolResultCharacters)
                    content = content.Substring(0, MaxToolResultCharacters) +
                               "\n…(工具结果过长已截断,关键信息保留在前段)";
                return new
                {
                    role = "tool",
                    tool_call_id = turn.ToolCallId,
                    content
                };
            }

            return new
            {
                role = string.IsNullOrWhiteSpace(turn.Role) ? "assistant" : turn.Role,
                content = turn.Content ?? string.Empty
            };
        }

        private List<object> BuildMessages(
            string userMessage,
            ExcelContextSnapshot excelContext,
            IReadOnlyList<ChatTurn> history)
        {
            var messages = new List<object>
            {
                new { role = "system", content = PromptBuilder.BuildSystemPrompt(_settings) }
            };
            if (excelContext != null)
                messages.Add(new { role = "system", content = excelContext.ToPromptText() });

            if (history != null)
                foreach (var turn in TrimHistoryToBudget(history))
                    messages.Add(ToProtocolMessage(turn));

            // userMessage 为空表示工具结果后的自动续跑；history 末尾已经是 tool。
            if (!string.IsNullOrWhiteSpace(userMessage))
                messages.Add(new { role = "user", content = userMessage });
            return messages;
        }

        /// <summary>
        /// 发送历史的字符预算。长任务中工具结果(含表格 JSON)会不断累积,
        /// 不裁剪会每轮全量重发,最终触发服务商上下文超限且该会话永久不可用。
        /// </summary>
        private const int MaxHistoryCharacters = 120000;

        /// <summary>单条工具结果参与发送的最大长度;超出部分截断。</summary>
        private const int MaxToolResultCharacters = 8000;

        private static IReadOnlyList<ChatTurn> TrimHistoryToBudget(IReadOnlyList<ChatTurn> history)
        {
            if (history == null || history.Count == 0) return history;
            var total = 0;
            for (var i = 0; i < history.Count; i++) total += TurnSize(history[i]);
            if (total <= MaxHistoryCharacters) return history;

            // 从末尾往前保留,超预算时回退到一条 user 消息处切割 ——
            // 保证保留部分仍以 user 开头,其后的 assistant(tool_calls)+tool 序列完整,
            // 不会产生协议上非法的"孤儿 tool 消息"。
            var kept = 0;
            for (var i = history.Count - 1; i >= 0; i--)
            {
                var size = TurnSize(history[i]);
                if (kept + size > MaxHistoryCharacters)
                {
                    var cut = i + 1;
                    while (cut < history.Count &&
                           !string.Equals(history[cut].Role, "user", StringComparison.OrdinalIgnoreCase))
                        cut++;

                    if (cut >= history.Count)
                    {
                        // 退化场景:整个 history 只有最初一条 user(Agent 自动续跑的典型形态),
                        // 没有 user 边界可切。保留尾部并补一条合成 user 头,保证协议合法
                        // 的同时把请求压进预算 —— 比全量发送触发服务商上下文超限好得多。
                        return TrimHistoryTail(history);
                    }

                    ThisAddIn.Log("历史裁剪: 发送时保留最近 " + (history.Count - cut) + "/" + history.Count +
                                  " 条消息(其余已超出上下文预算)");
                    var result = new List<ChatTurn>(history.Count - cut);
                    for (var index = cut; index < history.Count; index++) result.Add(history[index]);
                    return result;
                }
                kept += size;
            }
            return history;
        }

        /// <summary>
        /// 无 user 边界时的兜底:从尾部保留直到预算用尽,再在头部插入一条合成 user 消息
        /// 说明被省略的部分。切点同样避开"孤儿 tool"(若尾部第一组是 tool 开头,继续前移)。
        /// </summary>
        private static IReadOnlyList<ChatTurn> TrimHistoryTail(IReadOnlyList<ChatTurn> history)
        {
            var kept = 0;
            var cut = history.Count;
            for (var i = history.Count - 1; i >= 1; i--)
            {
                var size = TurnSize(history[i]);
                if (kept + size > MaxHistoryCharacters) break;
                kept += size;
                cut = i;
            }
            // 头部不能是孤儿 tool 消息:持续前移直到不是 tool。
            while (cut < history.Count &&
                   string.Equals(history[cut].Role, "tool", StringComparison.OrdinalIgnoreCase))
                cut++;
            if (cut <= 0 || cut >= history.Count) return history;

            var result = new List<ChatTurn>(history.Count - cut + 1)
            {
                new ChatTurn
                {
                    Role = "user",
                    Content = "(系统提示:为控制上下文长度,更早的对话与中间工具结果已省略," +
                              "以下是本任务最近的部分记录,请基于它继续。)"
                }
            };
            for (var index = cut; index < history.Count; index++) result.Add(history[index]);
            ThisAddIn.Log("历史裁剪(无边界兜底): 保留尾部 " + (history.Count - cut) + "/" + history.Count + " 条消息");
            return result;
        }

        /// <summary>按"实际发送口径"计算消息大小:工具结果超长部分发送前会截断,不计入预算。</summary>
        private static int TurnSize(ChatTurn turn)
        {
            if (turn == null) return 0;
            var contentLength = (turn.Content ?? string.Empty).Length;
            if (string.Equals(turn.Role, "tool", StringComparison.OrdinalIgnoreCase))
                contentLength = Math.Min(contentLength, MaxToolResultCharacters);
            var size = contentLength;
            if (turn.ToolCalls != null)
                foreach (var call in turn.ToolCalls)
                    size += (call?.ArgumentsJson ?? string.Empty).Length + 64;
            return size;
        }

        private static readonly JsonSerializerOptions JsonOpts = new JsonSerializerOptions
        {
            PropertyNamingPolicy = null,   // 保持 PascalCase 与匿名对象字段名一致(协议层会序列化为原样)
            WriteIndented = false
        };

        public void Dispose() { /* Http 为静态共享,不在此释放 */ }
    }
}
