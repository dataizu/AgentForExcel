namespace AgentForExcel.Models
{
    /// <summary>面向用户的能力入口。工具实现仍由 OperationDispatcher 管理。</summary>
    public sealed class CapabilityDefinition
    {
        public string Id { get; set; }
        public string Category { get; set; }
        public string Title { get; set; }
        public string Description { get; set; }
        public string Prompt { get; set; }
        public string Badge { get; set; }
        public string Accent { get; set; }
    }
}
