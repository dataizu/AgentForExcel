namespace AgentForExcel.Models
{
    /// <summary>
    /// A build-time product profile. This is intentionally not a subscription or
    /// payment implementation: commercial licensing is added only after the
    /// sellable packages have passed real customer validation.
    /// </summary>
    public enum ProductEdition
    {
        Trial = 0,
        Standard = 1,
        Professional = 2,
        Automation = 3
    }

    public static class ProductEditionInfo
    {
        public static ProductEdition Current
        {
            get
            {
#if AGENT_EDITION_TRIAL
                return ProductEdition.Trial;
#elif AGENT_EDITION_STANDARD
                return ProductEdition.Standard;
#elif AGENT_EDITION_AUTOMATION
                return ProductEdition.Automation;
#else
                return ProductEdition.Professional;
#endif
            }
        }

        public static string Id => Current.ToString();

        public static string DisplayName(ProductEdition edition)
        {
            switch (edition)
            {
                case ProductEdition.Trial: return "体验版";
                case ProductEdition.Standard: return "标准分析版";
                case ProductEdition.Automation: return "自动化交付版";
                default: return "专业自动化版";
            }
        }

        public static string CurrentDisplayName => DisplayName(Current);

        public static string Description(ProductEdition edition)
        {
            switch (edition)
            {
                case ProductEdition.Trial:
                    return "公式、图表和普通透视表体验；不含数据工程与自动化能力。";
                case ProductEdition.Standard:
                    return "包含数据体检、分析视图、公式、图表和普通透视表。";
                case ProductEdition.Automation:
                    return "在专业自动化版基础上提供受控 VBA 白名单配方。";
                default:
                    return "包含数据清洗、数据模型、看板与可重复分析流程；不含受控 VBA。";
            }
        }

        public static bool IsAtLeast(ProductEdition edition, ProductEdition minimumEdition)
        {
            return edition >= minimumEdition;
        }
    }
}
