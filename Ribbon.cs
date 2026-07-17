using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

namespace AgentForExcel
{
    /// <summary>
    /// 功能区 Ribbon。提供一个"显示/隐藏面板"按钮,
    /// 让用户能主动控制 AI 对话面板的可见性。
    /// </summary>
    [ComVisible(true)]
    public partial class Ribbon : Office.IRibbonExtensibility
    {
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
