namespace AgentForExcel.Operations
{
    // AgentLoopRunner 的独立回归测试只需要协议模型，不加载 Excel 执行层。
    public sealed class OperationCall
    {
        public string CallId { get; set; }
        public string ToolName { get; set; }
        public string ArgumentsJson { get; set; }
    }
}
