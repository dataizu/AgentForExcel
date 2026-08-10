using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using AgentForExcel.Models;
using AgentForExcel.Services;
using Office = Microsoft.Office.Core;

namespace AgentForExcel
{
    /// <summary>
    /// 功能区 Ribbon。按钮只负责把能力入口交给当前窗口的 ChatView；
    /// 能力提示词和版本边界统一由 CapabilityCatalog/EditionPolicy 提供。
    /// </summary>
    [ComVisible(true)]
    public partial class Ribbon : Office.IRibbonExtensibility
    {
        private static readonly IReadOnlyDictionary<string, string> CapabilityButtonMap =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "QualityButton", "quality" },
                { "TrendButton", "trend" },
                { "PivotButton", "pivot" },
                { "DashboardButton", "dashboard" },
                { "PowerQueryButton", "power-query" },
                { "RefreshButton", "refresh" },
                { "AutofitButton", "autofit" },
                { "ExportPdfButton", "export-pdf" },
                { "PowerPivotButton", "power-pivot" }
            };

        public Ribbon() { }

        /// <summary>加载 Ribbon XML(从嵌入资源/同目录文件读取)。</summary>
        public string GetCustomUI(string ribbonID)
        {
            var xml = GetResourceText("AgentForExcel.Ribbon.xml");
            ThisAddIn.Log("Ribbon.GetCustomUI: ribbonID=" + ribbonID + ", xmlLength=" + xml.Length);
            return xml;
        }

        /// <summary>Ribbon 加载回调。</summary>
        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            // 保留引用以便后续刷新(当前不需要)
            ThisAddIn.Log("Ribbon_Load: 功能区 XML 已加载");
        }

        /// <summary>"显示/隐藏面板"按钮点击:切换任务窗格可见性。</summary>
        public void ShowPanel_Click(Office.IRibbonControl control)
        {
            Globals.ThisAddIn.ToggleTaskPaneForActiveWindow();
        }

        /// <summary>
        /// 能力快捷入口：校验当前版本后，将目录中的完整提示词预填到当前窗口的
        /// ChatView 输入框并聚焦。这里不发送消息，也不直接执行 Excel 写操作。
        /// </summary>
        public void Capability_Click(Office.IRibbonControl control)
        {
            if (control == null || !CapabilityButtonMap.TryGetValue(control.Id, out var capabilityId))
            {
                ThisAddIn.Log("Ribbon: 未识别的能力按钮 " + (control?.Id ?? "<null>"));
                return;
            }

            var capability = CapabilityCatalog.Items.FirstOrDefault(item =>
                string.Equals(item.Id, capabilityId, System.StringComparison.OrdinalIgnoreCase));
            if (capability == null)
            {
                ThisAddIn.Log("Ribbon: 能力目录缺少 " + capabilityId);
                MessageBox.Show("未找到该能力的配置，请检查功能目录。", "Agent for Excel",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!EditionPolicy.IsCapabilityAvailable(capability.Id, ProductEditionInfo.Current))
            {
                var currentEdition = ProductEditionInfo.DisplayName(ProductEditionInfo.Current);
                var requiredEdition = ProductEditionInfo.DisplayName(capability.MinimumEdition);
                MessageBox.Show(
                    "当前版本为" + currentEdition + "；“" + capability.Title +
                    "”需要" + requiredEdition + "。请升级后使用。",
                    "Agent for Excel - 版本限制",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                ThisAddIn.Log("Ribbon: 能力不可用 " + capability.Id + "，当前版本=" + currentEdition);
                return;
            }

            Globals.ThisAddIn.PrefillCapabilityPrompt(capability.Id);
        }

        /// <summary>读取嵌入的资源文本(本程序集内、同目录的 Ribbon.xml)。</summary>
        private static string GetResourceText(string resourceName)
        {
            // 优先从程序集嵌入资源读
            var asm = Assembly.GetExecutingAssembly();
            using (var stream = asm.GetManifestResourceStream(resourceName))
            {
                if (stream != null)
                    using (var reader = new StreamReader(stream))
                        return reader.ReadToEnd();
            }
            // 兜底:从 DLL 同目录读文件
            var dir = Path.GetDirectoryName(asm.Location);
            var path = Path.Combine(dir, "Ribbon.xml");
            if (File.Exists(path))
                return File.ReadAllText(path);
            return string.Empty;
        }
    }
}
