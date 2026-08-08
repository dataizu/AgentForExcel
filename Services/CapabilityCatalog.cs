using System.Collections.Generic;
using AgentForExcel.Models;

namespace AgentForExcel.Services
{
    /// <summary>
    /// 功能中心的单一数据源。聊天欢迎页、功能中心和后续快捷入口均从这里读取，
    /// 避免在多个界面重复维护提示词和能力说明。
    /// </summary>
    public static class CapabilityCatalog
    {
        public static IReadOnlyList<CapabilityDefinition> Items { get; } = new[]
        {
            Item("quality", "分析数据", "数据质量体检", "检查字段类型、空值、重复、异常值和数据粒度，并给出处理优先级。",
                "请对当前工作表的完整数据区域执行数据质量体检：确认字段类型和数据粒度，检查空值、重复值、格式不一致、零值和异常值，并按影响程度给出处理优先级。", "只读分析", "#17734A"),
            Item("trend", "分析数据", "趋势与异常", "识别趋势、拐点和异常波动，并给出可追溯的原因线索。",
                "请分析当前工作表中的主要趋势、拐点和异常波动，说明时间范围、关键变化、影响最大的维度和可能原因。", "只读分析", "#17734A"),
            Item("compare", "分析数据", "分组对比", "比较区域、产品、渠道或时间段差异，定位差距最大的部分。",
                "请识别当前工作表中适合分组的字段，对核心数值指标做分组对比，并指出差异最大、最值得进一步核查的部分。", "只读分析", "#17734A"),
            Item("driver-tree", "分析数据", "指标树拆解", "把结果指标拆成可解释、可行动的驱动因素。",
                "请用指标树方法分析当前工作表：明确结果指标，拆解主要驱动因素，量化各因素影响并标出可直接行动的环节。", "分析方法", "#17734A"),
            Item("drilldown", "分析数据", "异常下钻", "从异常点逐层下钻到具体维度和记录。",
                "请从当前工作表识别异常点，并按时间、分类和业务维度逐层下钻，列出原因线索和需要进一步核查的记录范围。", "分析方法", "#17734A"),

            Item("formula", "生成报告", "生成公式", "根据业务目标生成可复制的 Excel 公式并解释参数。",
                "请根据当前工作表结构和我接下来描述的业务目标，生成可直接填充的 Excel 公式，说明引用范围、参数和适用条件；写入前先展示计划。", "需确认", "#B9652E"),
            Item("chart", "生成报告", "报告级图表", "自动选择横纵轴、聚合和降维规则，生成不拥挤的精美图表。",
                "请先对当前工作表完整数据区域做数据体检，识别时间、分类维度、数值指标、异常和重复粒度；再自动聚合重复横轴，分类过多时使用 Top-N+其他，生成可直接用于报告的精美图表，并核验标签与口径。", "新工作表", "#B9652E"),
            Item("pivot", "生成报告", "数据透视表", "根据字段角色配置行、列、筛选、值和聚合方式。",
                "请先识别当前工作表字段和数据粒度，再选择合适的行、列、筛选和值字段，在新的工作表中创建数据透视表并回读核验。", "新工作表", "#B9652E"),
            Item("dashboard", "生成报告", "联动数据看板", "生成 KPI、趋势、排名、占比、明细和联动筛选。",
                "请先对当前工作表做完整数据体检，确认核心指标、时间字段、分类维度和筛选维度，再新建一个保持源数据不变的联动数据看板，包含 KPI、趋势、排名、占比和明细，并核验筛选联动。", "新工作表", "#B9652E"),
            Item("analysis-view", "生成报告", "安全分析视图", "复制为值快照，在新工作表排序、筛选和展示。",
                "请基于当前工作表创建安全分析视图：保留源表不变，把数据复制为值快照，在新工作表完成需要的排序、筛选和报告级格式。", "保护源表", "#B9652E"),

            Item("power-query", "数据工程", "Power Query 清洗", "构建可刷新的去空、去重、类型转换、重命名和选列流程。",
                "请先体检当前工作表字段和数据质量，再创建 Power Query 清洗流程，处理空行、文本空格、重复、字段类型、必要的重命名和选列；加载到新工作表并回读核验，保持源数据不变。", "可刷新", "#356FA5"),
            Item("power-pivot", "数据工程", "Power Pivot 与 DAX", "载入数据模型、建立关系、创建度量值和模型透视表。",
                "请先检查当前 Power Pivot 数据模型和已有 Power Query，根据业务问题规划事实表、维度表和关系；需要时载入模型、创建 DAX 度量值，并在新工作表生成模型透视表后核验。", "专业能力", "#356FA5"),

            Item("refresh", "自动化", "刷新全部", "预览并执行受控 VBA，刷新连接、查询、模型和透视表。",
                "请先运行环境自检，然后预览受控 VBA 配方 refresh_all；说明影响并获得确认后执行，最后核验查询、模型和透视表刷新状态。", "受控 VBA", "#7B5AA6"),
            Item("autofit", "自动化", "统一自动列宽", "对可见工作表的已用区域执行受控自动列宽。",
                "请预览受控 VBA 配方 autofit_used_ranges，说明影响范围，获得确认后执行并核验主要工作表列宽。", "受控 VBA", "#7B5AA6"),
            Item("export-pdf", "自动化", "导出当前表 PDF", "通过白名单配方把当前工作表导出为 PDF。",
                "请确认当前活动工作表和导出路径，预览受控 VBA 配方 export_active_sheet_pdf，获得确认后执行并报告输出文件。", "受控 VBA", "#7B5AA6")
        };

        private static CapabilityDefinition Item(
            string id, string category, string title, string description, string prompt, string badge, string accent)
        {
            return new CapabilityDefinition
            {
                Id = id,
                Category = category,
                Title = title,
                Description = description,
                Prompt = prompt,
                Badge = badge,
                Accent = accent,
                MinimumEdition = EditionPolicy.GetMinimumEditionForCapability(id)
            };
        }
    }
}
