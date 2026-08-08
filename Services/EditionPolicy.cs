using System;
using AgentForExcel.Models;

namespace AgentForExcel.Services
{
    /// <summary>
    /// Keeps feature presentation and executable tool boundaries aligned for
    /// the build-time product edition. This is a product boundary, not DRM.
    /// </summary>
    public static class EditionPolicy
    {
        public static ProductEdition GetMinimumEditionForTool(string toolName)
        {
            if (string.IsNullOrWhiteSpace(toolName)) return ProductEdition.Trial;
            if (toolName.StartsWith("vba_", StringComparison.OrdinalIgnoreCase)) return ProductEdition.Automation;
            if (toolName.StartsWith("pq_", StringComparison.OrdinalIgnoreCase) ||
                toolName.StartsWith("pp_", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(toolName, "dashboard_create", StringComparison.OrdinalIgnoreCase))
                return ProductEdition.Professional;
            if (string.Equals(toolName, "data_profile", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(toolName, "analysis_create_view", StringComparison.OrdinalIgnoreCase))
                return ProductEdition.Standard;
            return ProductEdition.Trial;
        }

        public static ProductEdition GetMinimumEditionForCapability(string capabilityId)
        {
            switch (capabilityId ?? string.Empty)
            {
                case "quality":
                case "trend":
                case "compare":
                case "driver-tree":
                case "drilldown":
                case "analysis-view":
                    return ProductEdition.Standard;
                case "dashboard":
                case "power-query":
                case "power-pivot":
                    return ProductEdition.Professional;
                case "refresh":
                case "autofit":
                case "export-pdf":
                    return ProductEdition.Automation;
                default:
                    return ProductEdition.Trial;
            }
        }

        public static bool IsToolAvailable(string toolName, ProductEdition edition)
        {
            return ProductEditionInfo.IsAtLeast(edition, GetMinimumEditionForTool(toolName));
        }

        public static bool IsCapabilityAvailable(string capabilityId, ProductEdition edition)
        {
            return ProductEditionInfo.IsAtLeast(edition, GetMinimumEditionForCapability(capabilityId));
        }

        public static string GetUnavailableReason(string toolName, ProductEdition edition)
        {
            var required = GetMinimumEditionForTool(toolName);
            return "当前为" + ProductEditionInfo.DisplayName(edition) + "；该能力需要" +
                   ProductEditionInfo.DisplayName(required) + "。";
        }
    }
}
