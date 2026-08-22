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

        /// <summary>诊断日志:写到 %LOCALAPPDATA%\AgentForExcel\startup.log,超过 5MB 自动轮转。</summary>
        internal static void Log(string message)
        {
            try
            {
                var path = System.IO.Path.Combine(
                    System.Environment.GetFolderPath(System.Environment.SpecialFolder.LocalApplicationData),
                    "AgentForExcel", "startup.log");
                System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(path));
                // 轮转:重度使用下日志只增不减会无限膨胀;超限时保留一份旧档便于排查。
                // 每 64 条检查一次大小即可,避免每条都做文件元数据查询。
                if ((_logWriteCount++ & 63) == 0 && System.IO.File.Exists(path))
                {
                    var info = new System.IO.FileInfo(path);
                    if (info.Exists && info.Length > 5 * 1024 * 1024)
                    {
                        var archive = path + ".old";
                        System.IO.File.Delete(archive);
                        System.IO.File.Move(path, archive);
                    }
                }
                var line = $"[{DateTime.Now:HH:mm:ss.fff}] {message}";
                System.IO.File.AppendAllText(path, line + System.Environment.NewLine);
            }
            catch { /* 日志失败不影响加载项 */ }
        }

        private static int _logWriteCount;

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            // 释放托管资源 / 断开事件
            Application.WindowActivate -= Application_WindowActivate;
            Application.SheetSelectionChange -= Application_SheetSelectionChange;
            Application.SheetActivate -= Application_SheetActivate;
            Application.WorkbookActivate -= Application_WorkbookActivate;
            Application.WorkbookBeforeClose -= Application_WorkbookBeforeClose;
            Operations.Dashboard.DashboardInteractionManager.Shutdown();
            // 逐个退订面板的事件再清空字典:仅 Clear() 会让每个 ChatView
            // 连同聊天数据挂在 Selection 服务上等 GC。
            foreach (var pane in _taskPanesByWindow.Values)
            {
                try { FindChatView(pane)?.Teardown(); } catch { }
            }
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
            if (_appContext == null)
            {
                MessageBox.Show(
                    "Agent 尚未完成初始化(启动可能出过错)，请重启 Excel 后重试。",
                    "Agent for Excel",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            try
            {
                var pane = EnsureTaskPaneForWindow(window, false);
                if (pane == null) return;
                pane.Visible = !pane.Visible;
                Log("Ribbon: 切换窗口 " + window.Hwnd + " 的面板 (Visible=" + pane.Visible + ")");
            }
            catch (Exception ex)
            {
                // 面板可能已随窗口关闭而失效;吞掉异常避免冒泡进 Excel。
                Log("Ribbon: 切换面板异常: " + ex.Message);
            }
        }

        /// <summary>
        /// 显示当前窗口的任务窗格，并把能力目录中的提示词预填到 ChatView 输入框。
        /// 该入口只准备对话内容，不发送消息，也不直接执行工作簿写操作。
        /// </summary>
        internal void PrefillCapabilityPrompt(string capabilityId)
        {
            if (string.IsNullOrWhiteSpace(capabilityId)) return;

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

            AgentForExcel.Models.CapabilityDefinition capability = null;
            foreach (var item in Services.CapabilityCatalog.Items)
            {
                if (string.Equals(item.Id, capabilityId, StringComparison.OrdinalIgnoreCase))
                {
                    capability = item;
                    break;
                }
            }

            if (capability == null)
            {
                Log("Ribbon: 预填时未找到能力目录项 " + capabilityId);
                MessageBox.Show(
                    "未找到该能力的提示词配置，请检查功能目录。",
                    "Agent for Excel",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            try
            {
                var pane = EnsureTaskPaneForWindow(window, false);
                pane.Visible = true;

                // 通过 CustomTaskPane -> TaskPaneHost -> ElementHost 找到当前窗口的 ChatView，
                // 保持多窗口之间的任务窗格和输入状态彼此独立。
                var chatView = FindChatView(pane);
                var inputBox = FindChatInputBox(chatView);
                if (inputBox == null)
                {
                    Log("Ribbon: 窗口 " + window.Hwnd + " 未找到 ChatView 输入框");
                    MessageBox.Show(
                        "AI 面板尚未准备好，请稍后重试。",
                        "Agent for Excel",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                    return;
                }

                System.Action applyPrompt = () =>
                {
                    inputBox.Text = capability.Prompt ?? string.Empty;
                    inputBox.CaretIndex = inputBox.Text.Length;
                    inputBox.SelectionLength = 0;
                    inputBox.Focus();
                    System.Windows.Input.Keyboard.Focus(inputBox);
                };

                if (inputBox.Dispatcher.CheckAccess())
                    applyPrompt();
                else
                    inputBox.Dispatcher.Invoke(applyPrompt);

                Log("Ribbon: 已向窗口 " + window.Hwnd + " 预填能力 " + capability.Id);
            }
            catch (Exception ex)
            {
                Log("Ribbon: 预填能力 " + capabilityId + " 异常: " + ex);
                MessageBox.Show(
                    "AI 面板暂时不可用，请稍后重试。",
                    "Agent for Excel",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
        }

        private static ChatView FindChatView(CustomTaskPane pane)
        {
            var host = pane?.Control as TaskPaneHost;
            if (host == null) return null;

            foreach (Control child in host.Controls)
            {
                var elementHost = child as ElementHost;
                if (elementHost?.Child is ChatView chatView)
                    return chatView;
            }

            return null;
        }

        private static System.Windows.Controls.TextBox FindChatInputBox(ChatView chatView)
        {
            if (chatView == null) return null;

            var namedInput = chatView.FindName("InputBox") as System.Windows.Controls.TextBox;
            if (namedInput != null) return namedInput;

            return FindChatInputBoxInVisualTree(chatView);
        }

        private static System.Windows.Controls.TextBox FindChatInputBoxInVisualTree(
            System.Windows.DependencyObject parent)
        {
            if (parent == null) return null;

            var textBox = parent as System.Windows.Controls.TextBox;
            var namedElement = parent as System.Windows.FrameworkElement;
            if (textBox != null && string.Equals(namedElement?.Name, "InputBox", StringComparison.Ordinal))
                return textBox;

            var childCount = System.Windows.Media.VisualTreeHelper.GetChildrenCount(parent);
            for (var index = 0; index < childCount; index++)
            {
                var result = FindChatInputBoxInVisualTree(
                    System.Windows.Media.VisualTreeHelper.GetChild(parent, index));
                if (result != null) return result;
            }

            return null;
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

            // 工作簿关闭即其窗口销毁:立即清理对应的任务窗格条目,
            // 不等下次激活时的惰性检测(死条目会钉住整棵 ChatView 对象图)。
            try
            {
                var doomedHandles = new List<int>();
                foreach (var pair in _taskPanesByWindow)
                {
                    try
                    {
                        var paneWindow = pair.Value.Window as Excel.Window;
                        var ownerWorkbookName = Convert.ToString((paneWindow?.Parent as Excel.Workbook)?.Name);
                        if (string.Equals(ownerWorkbookName, workbook?.Name, StringComparison.OrdinalIgnoreCase))
                            doomedHandles.Add(pair.Key);
                    }
                    catch
                    {
                        // 访问失败说明窗口已销毁,同样需要清理。
                        doomedHandles.Add(pair.Key);
                    }
                }
                foreach (var handle in doomedHandles)
                {
                    CustomTaskPane doomed;
                    if (_taskPanesByWindow.TryGetValue(handle, out doomed))
                    {
                        try { FindChatView(doomed)?.Teardown(); } catch { }
                    }
                    _taskPanesByWindow.Remove(handle);
                }
            }
            catch (Exception ex) { Log("WorkbookBeforeClose 清理面板异常: " + ex.Message); }
        }

        /// <summary>获取或创建指定 Excel 窗口独有的任务窗格。</summary>
        private CustomTaskPane EnsureTaskPaneForWindow(Excel.Window window, bool showWhenCreated)
        {
            if (window == null) throw new ArgumentNullException(nameof(window));
            if (_appContext == null)
            {
                // 启动半失败时 ChatView.Initialize(null) 会空引用;直接跳过。
                Log("EnsureTaskPaneForWindow: 上下文未初始化,跳过");
                return null;
            }

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
                catch (Exception)
                {
                    // RCW 释放后抛出的常是 COMException 而非 ObjectDisposedException,
                    // 任何失败都按"窗口已关闭"处理:退订事件并清理死条目。
                    try { FindChatView(existing)?.Teardown(); } catch { }
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
