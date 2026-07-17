namespace AgentForExcel.Operations
{
    /// <summary>
    /// 所有 Excel 操作的统一抽象。
    /// 每个 IOperation 实例代表 LLM 下发的一条操作指令,
    /// 由 OperationDispatcher 解析 LLM 返回的 JSON 后构造。
    /// </summary>
    public interface IOperation
    {
        /// <summary>工具名,如 "cell_read_range"、"pp_add_measure"。</summary>
        string ToolName { get; }

        /// <summary>是否有副作用(写、改、删)。用于决定是否需要用户确认。</summary>
        bool IsWriteOperation { get; }

        /// <summary>给用户看的人类可读描述(确认对话框用)。</summary>
        string Describe();

        /// <summary>在 Excel 中执行此操作,返回结果文本(展示给用户)。</summary>
        string Execute(AppContext context);
    }

    /// <summary>把 LLM 的 JSON 参数解析成 IOperation 的工厂。</summary>
    /// <remarks>
    /// 每个工具对应一个 Factory,Dispatcher 维护 工具名→工厂 的映射。
    /// 新增能力时只需:实现一个 Op + 一个 Factory,然后注册到 Dispatcher。
    /// </remarks>
    public interface IOperationFactory
    {
        string ToolName { get; }
        IOperation Parse(string argumentsJson);
    }
}
