using System;
using System.Collections.Generic;
using System.Text;
using System.Text.Json;
using AgentForExcel.Operations;

namespace AgentForExcel.AI
{
    /// <summary>累积 OpenAI 兼容 SSE 中被拆分的文本和 tool_calls delta。</summary>
    internal sealed class StreamingReplyAccumulator
    {
        private readonly StringBuilder _text = new StringBuilder();
        private readonly SortedDictionary<int, ToolCallAccumulator> _tools =
            new SortedDictionary<int, ToolCallAccumulator>();

        public string Consume(string jsonChunk)
        {
            if (string.IsNullOrWhiteSpace(jsonChunk)) return string.Empty;

            using (var doc = JsonDocument.Parse(jsonChunk))
            {
                var root = doc.RootElement;
                if (!root.TryGetProperty("choices", out var choices) ||
                    choices.ValueKind != JsonValueKind.Array || choices.GetArrayLength() == 0)
                    return string.Empty;

                var choice = choices[0];
                if (!choice.TryGetProperty("delta", out var delta) || delta.ValueKind != JsonValueKind.Object)
                {
                    // 少数兼容接口即使收到 stream=true 仍返回普通 message JSON。
                    if (!choice.TryGetProperty("message", out var message) || message.ValueKind != JsonValueKind.Object)
                        return string.Empty;
                    return ConsumeCompleteMessage(message);
                }

                var textDelta = string.Empty;
                if (delta.TryGetProperty("content", out var content) && content.ValueKind == JsonValueKind.String)
                {
                    textDelta = content.GetString() ?? string.Empty;
                    _text.Append(textDelta);
                }

                if (delta.TryGetProperty("tool_calls", out var toolCalls) && toolCalls.ValueKind == JsonValueKind.Array)
                {
                    var ordinal = 0;
                    foreach (var toolCall in toolCalls.EnumerateArray())
                    {
                        var index = toolCall.TryGetProperty("index", out var indexElement) &&
                                    indexElement.ValueKind == JsonValueKind.Number
                            ? indexElement.GetInt32()
                            : ordinal;
                        ordinal++;

                        if (!_tools.TryGetValue(index, out var target))
                        {
                            target = new ToolCallAccumulator();
                            _tools[index] = target;
                        }

                        if (toolCall.TryGetProperty("id", out var id) && id.ValueKind == JsonValueKind.String)
                            target.Id = id.GetString() ?? target.Id;

                        if (!toolCall.TryGetProperty("function", out var function) ||
                            function.ValueKind != JsonValueKind.Object) continue;

                        if (function.TryGetProperty("name", out var name) && name.ValueKind == JsonValueKind.String)
                            target.Name.Append(name.GetString());
                        if (function.TryGetProperty("arguments", out var arguments) && arguments.ValueKind == JsonValueKind.String)
                            target.Arguments.Append(arguments.GetString());
                    }
                }

                return textDelta;
            }
        }

        private string ConsumeCompleteMessage(JsonElement message)
        {
            var text = string.Empty;
            if (message.TryGetProperty("content", out var content) && content.ValueKind == JsonValueKind.String)
            {
                text = content.GetString() ?? string.Empty;
                _text.Append(text);
            }

            if (!message.TryGetProperty("tool_calls", out var toolCalls) || toolCalls.ValueKind != JsonValueKind.Array)
                return text;

            var ordinal = 0;
            foreach (var toolCall in toolCalls.EnumerateArray())
            {
                var index = toolCall.TryGetProperty("index", out var indexElement) &&
                            indexElement.ValueKind == JsonValueKind.Number
                    ? indexElement.GetInt32()
                    : ordinal;
                ordinal++;
                var target = new ToolCallAccumulator();
                if (toolCall.TryGetProperty("id", out var id) && id.ValueKind == JsonValueKind.String)
                    target.Id = id.GetString();
                if (toolCall.TryGetProperty("function", out var function) && function.ValueKind == JsonValueKind.Object)
                {
                    if (function.TryGetProperty("name", out var name) && name.ValueKind == JsonValueKind.String)
                        target.Name.Append(name.GetString());
                    if (function.TryGetProperty("arguments", out var arguments) && arguments.ValueKind == JsonValueKind.String)
                        target.Arguments.Append(arguments.GetString());
                }
                _tools[index] = target;
            }
            return text;
        }

        public LlmReply BuildReply()
        {
            var operations = new List<OperationCall>();
            foreach (var item in _tools)
            {
                operations.Add(new OperationCall
                {
                    CallId = item.Value.Id,
                    ToolName = item.Value.Name.ToString(),
                    ArgumentsJson = item.Value.Arguments.Length == 0 ? "{}" : item.Value.Arguments.ToString()
                });
            }

            return new LlmReply { Text = _text.ToString(), Operations = operations };
        }

        private sealed class ToolCallAccumulator
        {
            public string Id { get; set; }
            public StringBuilder Name { get; } = new StringBuilder();
            public StringBuilder Arguments { get; } = new StringBuilder();
        }
    }
}
