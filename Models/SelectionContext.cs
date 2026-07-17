using System;

namespace AgentForExcel.Models
{
    /// <summary>当前或已锁定的 Excel 选区元数据，不持有 COM 对象。</summary>
    public sealed class SelectionContext
    {
        public string WorkbookName { get; set; }
        public string WorkbookFullName { get; set; }
        public string SheetName { get; set; }
        public string Address { get; set; }
        public int RowCount { get; set; }
        public int ColumnCount { get; set; }
        public long CellCount { get; set; }
        public bool IsMultiArea { get; set; }
        public DateTime CapturedAtUtc { get; set; } = DateTime.UtcNow;

        public bool IsValid => !string.IsNullOrWhiteSpace(WorkbookName) &&
                               !string.IsNullOrWhiteSpace(SheetName) &&
                               !string.IsNullOrWhiteSpace(Address);

        public string DisplayText
        {
            get
            {
                if (!IsValid) return "未检测到 Excel 选区";
                var size = IsMultiArea
                    ? $"多区域 · {CellCount:#,##0} 个单元格"
                    : $"{RowCount:#,##0} 行 × {ColumnCount:#,##0} 列";
                return $"{SheetName}!{Address} · {size}";
            }
        }

        public string PromptReference => IsValid
            ? $"工作簿「{WorkbookName}」/ 工作表「{SheetName}」/ 区域 {Address}（{RowCount} 行 × {ColumnCount} 列）"
            : "当前没有有效选区";

        public SelectionContext Clone()
        {
            return new SelectionContext
            {
                WorkbookName = WorkbookName,
                WorkbookFullName = WorkbookFullName,
                SheetName = SheetName,
                Address = Address,
                RowCount = RowCount,
                ColumnCount = ColumnCount,
                CellCount = CellCount,
                IsMultiArea = IsMultiArea,
                CapturedAtUtc = CapturedAtUtc
            };
        }
    }
}
