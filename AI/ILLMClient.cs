using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using AgentForExcel.Models;
using AgentForExcel.Operations;

namespace AgentForExcel.AI
{
    /// <summary>
    /// AI 客户端统一接口。无论后端是 GLM / DeepSeek / 通义 / OpenAI,UI 层只面向此接口。
    /// </summary>
    public interface ILLMClient
    {
        /// <summary>
        /// 发送一轮对话,返回助手回复(自然语言文本 + 可能的操作指令)。
        /// </summary>
        /// <param name="userMessage">用户这一轮的输入。</param>
        /// <param name="excelContext">当前 Excel 上下文快照(作为系统上下文)。</param>
        Task<LlmReply> ChatAsync(string userMessage, ExcelContextSnapshot excelContext);

        /// <summary>多轮对话重载:带上历史消息。</summary>
        Task<LlmReply> ChatAsync(string userMessage, ExcelContextSnapshot excelContext, IReadOnlyList<ChatTurn> history);

        /// <summary>流式多轮对话；文本增量通过 onTextDelta 实时返回。</summary>
        Task<LlmReply> ChatStreamingAsync(
            string userMessage,
            ExcelContextSnapshot excelContext,
            IReadOnlyList<ChatTurn> history,
            Action<string> onTextDelta);
    }

}
