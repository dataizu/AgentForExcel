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
                        $"大模型请求失败 ({(int)resp.StatusCode} {resp.StatusCode}):\n{respText}");
                return ParseReply(respText);
            }
        }

        public async Task<LlmReply> ChatStreamingAsync(
            string userMessage,
            ExcelContextSnapshot excelContext,
            IReadOnlyList<ChatTurn> history,
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
            for (var attempt = 0; attempt < 2; attempt++)
            {
                try
                {
                    return await SendStreamingOnceAsync(
                        url,
                        jsonBody,
                        onTextDelta,
                        forceNewConnection: attempt > 0);
                }
                catch (TaskCanceledException ex)
                {
                    throw new InvalidOperationException("模型请求超过 180 秒，已停止等待。请缩小分析范围后重试。", ex);
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

                if (attempt == 0)
                    await Task.Delay(400);
            }

            var host = new Uri(url).Host;
            throw new InvalidOperationException(
                $"无法连接模型服务 {host}，自动重试仍失败：{GetDeepestMessage(lastConnectionError)}",
                lastConnectionError);
        }

        private async Task<LlmReply> SendStreamingOnceAsync(
            string url,
            string jsonBody,
            Action<string> onTextDelta,
            bool forceNewConnection)
        {
            using (var request = new HttpRequestMessage(HttpMethod.Post, url))
            {
                request.Headers.Add("Authorization", "Bearer " + _settings.ApiKey);
                request.Headers.Add("Accept", "text/event-stream");
                if (forceNewConnection) request.Headers.ConnectionClose = true;
                request.Content = new StringContent(jsonBody, Encoding.UTF8, "application/json");

                using (var response = await Http.SendAsync(request, HttpCompletionOption.ResponseHeadersRead))
                {
                    if (!response.IsSuccessStatusCode)
                    {
                        var error = await response.Content.ReadAsStringAsync();
                        throw new InvalidOperationException(
                            $"大模型请求失败 ({(int)response.StatusCode} {response.StatusCode}):\n{error}");
                    }

                    var accumulator = new StreamingReplyAccumulator();
                    var receivedChunk = false;
                    using (var stream = await response.Content.ReadAsStreamAsync())
                    using (var reader = new StreamReader(stream, Encoding.UTF8))
                    {
                        string line;
                        while ((line = await reader.ReadLineAsync()) != null)
                        {
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
                    }

                    if (!receivedChunk)
                        throw new IOException("模型连接已建立，但没有收到可识别的流式数据。");
                    var reply = accumulator.BuildReply();
                    ThisAddIn.Log($"LLM 流式完成: textLength={reply.Text?.Length ?? 0}, toolCalls={reply.Operations.Count}");
                    return reply;
                }
            }
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
                return new
                {
                    role = "tool",
                    tool_call_id = turn.ToolCallId,
                    content = turn.Content ?? string.Empty
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
                foreach (var turn in history)
                    messages.Add(ToProtocolMessage(turn));

            // userMessage 为空表示工具结果后的自动续跑；history 末尾已经是 tool。
            if (!string.IsNullOrWhiteSpace(userMessage))
                messages.Add(new { role = "user", content = userMessage });
            return messages;
        }

        private static readonly JsonSerializerOptions JsonOpts = new JsonSerializerOptions
        {
            PropertyNamingPolicy = null,   // 保持 PascalCase 与匿名对象字段名一致(协议层会序列化为原样)
            WriteIndented = false
        };

        public void Dispose() { /* Http 为静态共享,不在此释放 */ }
    }
}
