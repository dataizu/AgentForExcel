using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using AgentForExcel.UI;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools;

namespace AgentForExcel
{
    /// <summary>
    /// VSTO 加载项入口。Excel 启动时实例化,在此挂载右侧 AI 对话面板。
    /// </summary>
    public partial class ThisAddIn
    {
        // Excel 2013+ 的工作簿窗口彼此独立。每个窗口需要自己的任务窗格和 UI 控件。
        private readonly Dictionary<int, CustomTaskPane> _taskPanesByWindow =
            new Dictionary<int, CustomTaskPane>();
        // 全局应用上下文(各操作模块通过它访问 Excel 对象模型)
        private AppContext _appContext;
        private Services.AgentAutomationService _automationService;

        /// <summary>提供对当前应用上下文的静态访问(供 UI/操作层使用)。</summary>
        internal static AppContext App { get; private set; }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            Log("===== AgentForExcel 启动 =====");
            try
            {
                // 0) 确保 WPF Application 上下文存在(VSTO 宿主默认没有,
                //    缺失会导致 WPF 控件主题/资源加载异常)
                Log("步骤0: 初始化 WPF Application...");
                EnsureWpfApplication();
                Log("步骤0: 完成");

                // 1) 初始化应用上下文 —— 持有 Excel Application 与服务容器
                Log("步骤1: 创建 AppContext...");
                _appContext = new AppContext(Application, Globals.Factory);
                App = _appContext;
                Log("步骤1: AppContext.Initialize()...");
                _appContext.Initialize();
                Operations.Dashboard.DashboardInteractionManager.Initialize(Application);
                Log("步骤1: 完成");

                // 2) 为当前窗口创建真正的 WPF 对话面板；后续窗口激活时按窗口补建。
                Log("步骤2: 注册 Excel 窗口事件...");
                Application.WindowActivate += Application_WindowActivate;
                Application.SheetSelectionChange += Application_SheetSelectionChange;
                Application.SheetActivate += Application_SheetActivate;
                Application.WorkbookActivate += Application_WorkbookActivate;
                Application.WorkbookBeforeClose += Application_WorkbookBeforeClose;
                if (Application.ActiveWindow != null)
                    EnsureTaskPaneForWindow(Application.ActiveWindow, true);
                Log("步骤2: 完成");

                Log("===== 启动成功 =====");
            }
            catch (Exception ex)
            {
                Log("!!!! 启动异常 !!!!");
                Log("类型: " + ex.GetType().FullName);
                Log("消息: " + ex.Message);
                Log("堆栈: " + ex.StackTrace);
                if (ex.InnerException != null)
                {
                    Log("内部异常: " + ex.InnerException.Message);
                    Log("内部堆栈: " + ex.InnerException.StackTrace);
                }
                Log("!!!! 异常结束 !!!!");
            }
        }

        /// <summary>诊断日志:写到 %LOCALAPPDATA%\AgentForExcel\startup.log</summary>
        internal static void Log(string message)
        {
            try
            {
                var path = System.IO.Path.Combine(
                    System.Environment.GetFolderPath(System.Environment.SpecialFolder.LocalApplicationData),
                    "AgentForExcel", "startup.log");
                System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(path));
                var line = $"[{DateTime.Now:HH:mm:ss.fff}] {message}";
                System.IO.File.AppendAllText(path, line + System.Environment.NewLine);
            }
            catch { /* 日志失败不影响加载项 */ }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            // 释放托管资源 / 断开事件
            Application.WindowActivate -= Application_WindowActivate;
            Application.SheetSelectionChange -= Application_SheetSelectionChange;
            Application.SheetActivate -= Application_SheetActivate;
            Application.WorkbookActivate -= Application_WorkbookActivate;
            Application.WorkbookBeforeClose -= Application_WorkbookBeforeClose;
            Operations.Dashboard.DashboardInteractionManager.Shutdown();
            _taskPanesByWindow.Clear();
            _appContext?.Dispose();
        }

        /// <summary>
        /// 向 Office 提供 Ribbon XML 扩展。VSTO 会在加载功能区时调用此方法。
        /// </summary>
        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            Log("CreateRibbonExtensibilityObject: 提供 Ribbon");
            return _ribbon ?? (_ribbon = new Ribbon());
        }

        /// <summary>向 COMAddIn.Object 暴露只读诊断接口，供安装验收与运维检查使用。</summary>
        protected override object RequestComAddInAutomationService()
        {
            return _automationService ?? (_automationService = new Services.AgentAutomationService());
        }

        private Ribbon _ribbon;

        /// <summary>切换当前 Excel 窗口对应的任务窗格。</summary>
        internal void ToggleTaskPaneForActiveWindow()
        {
            var window = Application.ActiveWindow;
            if (window == null)
            {
                MessageBox.Show(
                    "当前没有可用的 Excel 工作簿窗口。",
                    "Agent for Excel",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            var pane = EnsureTaskPaneForWindow(window, false);
            pane.Visible = !pane.Visible;
            Log("Ribbon: 切换窗口 " + window.Hwnd + " 的面板 (Visible=" + pane.Visible + ")");
        }

        private void Application_WindowActivate(Excel.Workbook workbook, Excel.Window window)
        {
            try
            {
                _appContext?.Selection?.Refresh();
                EnsureTaskPaneForWindow(window, true);
            }
            catch (Exception ex)
            {
                Log("WindowActivate 创建面板异常: " + ex);
            }
        }

        private void Application_SheetSelectionChange(object sheet, Excel.Range target)
        {
            try
            {
                _appContext?.Selection?.Update(sheet as Excel.Worksheet, target);
            }
            catch (Exception ex)
            {
                Log("SheetSelectionChange 更新选区异常: " + ex.Message);
            }
        }

        private void Application_SheetActivate(object sheet)
        {
            try { _appContext?.Selection?.Refresh(); }
            catch (Exception ex) { Log("SheetActivate 更新选区异常: " + ex.Message); }
        }

        private void Application_WorkbookActivate(Excel.Workbook workbook)
        {
            try { _appContext?.Selection?.Refresh(); }
            catch (Exception ex) { Log("WorkbookActivate 更新选区异常: " + ex.Message); }
        }

        private void Application_WorkbookBeforeClose(Excel.Workbook workbook, ref bool cancel)
        {
            if (cancel) return;
            try
            {
                var locked = _appContext?.Selection?.Locked;
                if (locked != null && string.Equals(locked.WorkbookName, workbook?.Name, StringComparison.OrdinalIgnoreCase))
                    _appContext.Selection.Unlock();
            }
            catch (Exception ex) { Log("WorkbookBeforeClose 清理选区异常: " + ex.Message); }
        }

        /// <summary>获取或创建指定 Excel 窗口独有的任务窗格。</summary>
        private CustomTaskPane EnsureTaskPaneForWindow(Excel.Window window, bool showWhenCreated)
        {
            if (window == null) throw new ArgumentNullException(nameof(window));

            var windowHandle = window.Hwnd;
            CustomTaskPane existing;
            if (_taskPanesByWindow.TryGetValue(windowHandle, out existing))
            {
                try
                {
                    // 访问属性可以识别已随窗口关闭而释放的旧窗格。
                    var ignored = existing.Window;
                    return existing;
                }
                catch (ObjectDisposedException)
                {
                    _taskPanesByWindow.Remove(windowHandle);
                }
            }

            Log("创建窗口 " + windowHandle + " 的 ChatView...");
            var chatView = new ChatView();
            chatView.Initialize(_appContext);

            var elementHost = new ElementHost
            {
                Dock = DockStyle.Fill,
                Child = chatView
            };
            var host = new TaskPaneHost { Dock = DockStyle.Fill };
            host.Controls.Add(elementHost);

            var pane = CustomTaskPanes.Add(host, "Agent for Excel", window);
            pane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
            pane.Width = 380;
            pane.Visible = showWhenCreated;

            _taskPanesByWindow[windowHandle] = pane;
            Log("窗口 " + windowHandle + " 的面板创建完成 (Visible=" + pane.Visible + ")");
            return pane;
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }

        #endregion

        /// <summary>
        /// 确保 WPF 的 Application 实例存在。
        /// VSTO/Office 宿主默认不启动 WPF Application,会导致
        /// Application.Current 为 null、主题资源缺失、控件显示异常。
        /// 这里在不存在时创建一个隐藏实例,加载默认主题。
        /// </summary>
        private static void EnsureWpfApplication()
        {
            if (System.Windows.Application.Current != null) return;
            try
            {
                // 创建隐藏的 WPF Application 实例(不显示窗口)
                new System.Windows.Application
                {
                    ShutdownMode = System.Windows.ShutdownMode.OnExplicitShutdown
                };
            }
            catch { /* 已存在或创建失败时忽略,不阻塞加载项启动 */ }
        }
    }
}
