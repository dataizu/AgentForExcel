using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Threading;
using AgentForExcel.AI;
using AgentForExcel.Models;
using AgentForExcel.Operations;
using AgentForExcel.Operations.Analysis;
using AgentForExcel.Operations.Cell;
using AgentForExcel.Operations.Tasking;
using AgentForExcel.Services;

namespace AgentForExcel.UI
{
    /// <summary>
    /// AI 对话面板。负责:渲染消息流、接收用户输入、调用 LLM、
    /// 把 LLM 返回的操作指令交给 OperationDispatcher(阶段 2 起生效)。
    /// </summary>
    public partial class ChatView : UserControl
    {
        private const int StreamingRenderIntervalMilliseconds = 120;
        private AppContext _app;
        private readonly ObservableCollection<ChatMessage> _messages = new ObservableCollection<ChatMessage>();
        private readonly List<ChatTurn> _history = new List<ChatTurn>();
        private ChatHistoryStore _chatStore;
        private ChatHistoryDocument _chatHistory;
        private ChatConversation _activeConversation;
        private bool _updatingModelPicker;
        private bool _isBusy;
        private bool _scrollToBottomPending;

        public ChatView()
        {
            ThisAddIn.Log("ChatView: 构造开始");
            try
            {
                InitializeComponent();
                ThisAddIn.Log("ChatView: InitializeComponent 完成");
                MessagesList.ItemsSource = _messages;
                ThisAddIn.Log("ChatView: ItemsSource 已设置");
            }
            catch (Exception ex)
            {
                ThisAddIn.Log("ChatView 构造异常: " + ex);
                throw;
            }
        }

        /// <summary>由 ThisAddIn 在启动时注入应用上下文。</summary>
        public void Initialize(AppContext app)
        {
            ThisAddIn.Log("ChatView.Initialize: 开始");
            _app = app;
            _chatStore = new ChatHistoryStore();
            _chatHistory = _app.Settings.SaveChatHistory
                ? _chatStore.Load()
                : new ChatHistoryDocument();
            if (_chatHistory.Conversations.Count == 0)
                _chatStore.CreateConversation(_chatHistory, _app.Settings.ActiveProfileId);
            var activeConversation = _chatHistory.Conversations.FirstOrDefault(item =>
                                         item.Id == _chatHistory.ActiveConversationId)
                                     ?? _chatHistory.Conversations[0];
            ActivateConversation(activeConversation, false);
            RefreshModelPicker();
            UpdateActiveModelLabel();
            if (_app.Selection != null)
            {
                _app.Selection.Changed += SelectionContext_Changed;
                _app.Selection.Refresh();
                UpdateSelectionContextUi();
            }
            ThisAddIn.Log("ChatView.Initialize: 完成,消息数=" + _messages.Count);
        }

        private void ChatView_OnLoaded(object sender, RoutedEventArgs e)
        {
            // Excel 任务窗格首次挂载时可能把滚动区域带到输入区附近，
            // 布局完成后显式回到欢迎页顶部，再把键盘焦点交给输入框。
            Dispatcher.BeginInvoke(new Action(() =>
            {
                if (_messages.Count == 0) MessagesScroll.ScrollToTop();
                else MessagesScroll.ScrollToBottom();
                InputBox.Focus();
            }));
        }

        private async void SendButton_Click(object sender, RoutedEventArgs e) => await SendAsync();

        private async void InputBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter && (Keyboard.Modifiers & ModifierKeys.Shift) == 0)
            {
                e.Handled = true;
                await SendAsync();
            }
        }

        private async Task SendAsync()
        {
            if (_app == null) return;

            var text = InputBox.Text?.Trim();
            if (string.IsNullOrEmpty(text)) return;

            if (!_app.Selection.IsLocked)
                _app.Selection.LockCurrent("task");
            var effectiveSelection = _app.Selection.Effective;
            if (text.IndexOf("@当前选区", StringComparison.OrdinalIgnoreCase) >= 0 &&
                (effectiveSelection == null || !effectiveSelection.IsValid))
            {
                MessageBox.Show("请先在 Excel 中选中要引用的区域。", "当前选区",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            var requestText = effectiveSelection != null
                ? text.Replace("@当前选区", effectiveSelection.PromptReference)
                : text;
            var preserveTaskSelectionLock = false;

            // 渲染用户消息并清空输入框
            WelcomePanel.Visibility = Visibility.Collapsed;
            AddUser(text);
            UpdateConversationTitle(text);
            PersistActiveConversation();
            InputBox.Clear();
            SetInputEnabled(false);

            var thinking = AddStatus("正在分析工作簿…");
            var chatTimer = Stopwatch.StartNew();
            var chatOutcome = "exception";
            var chatRounds = 0;
            var chatToolCalls = 0;
            var streamedCharacters = 0;
            var streamedRenderCount = 0;

            try
            {
                // 1) 校验配置
                if (string.IsNullOrWhiteSpace(_app.Settings.ApiKey))
                {
                    chatOutcome = "configuration_missing";
                    thinking.Kind = ChatMessageKind.Text;
                    thinking.Text = "尚未配置大模型 API Key。请打开右上角设置，在“模型与连接”中填写后再试。";
                    return;
                }

                // 2) 启动真正的多轮 Agent 循环：工具结果会按 tool_call_id 回传给模型，
                // 模型可以继续分析或调用下一项工具，直到给出不含工具调用的最终答复。
                ChatMessage currentStatus = thinking;
                ChatMessage reusableToolStatus = null;
                ChatMessage activeTaskPlan = null;
                var streamedText = new StringBuilder();
                var renderTimer = Stopwatch.StartNew();
                var run = await AgentLoopRunner.RunAsync(
                    requestText,
                    _history,
                    async turns =>
                    {
                        var contextSnapshot = ExcelContextSnapshot.Capture(_app.Excel, _app.Selection.Effective);
                        return await _app.LLM.ChatStreamingAsync(
                            null,
                            contextSnapshot,
                            turns,
                            delta =>
                            {
                                streamedText.Append(delta);
                                streamedCharacters += delta?.Length ?? 0;
                                // MarkdownViewer 每次更新都会重建完整的 WPF 视觉树；限频可避免
                                // 长回复把 UI 线程和滚动队列占满，最终答复仍会立即完整渲染。
                                if (renderTimer.ElapsedMilliseconds < StreamingRenderIntervalMilliseconds) return;
                                SetAssistantText(currentStatus, streamedText.ToString());
                                streamedRenderCount++;
                                renderTimer.Restart();
                            });
                    },
                    calls => _app.Dispatcher.ExecuteAsync(calls, OnConfirmOperation),
                    round =>
                    {
                        if (round > 1)
                        {
                            if (currentStatus.Kind == ChatMessageKind.Status)
                                SetStatus(currentStatus, "正在整理最终结果…");
                            else
                                currentStatus = AddStatus("正在根据工具结果继续分析…");
                        }
                        else
                            SetStatus(currentStatus, "正在分析工作簿…");
                        reusableToolStatus = null;
                        streamedText.Clear();
                        renderTimer.Restart();
                    },
                    (reply, round) =>
                    {
                        if (!string.IsNullOrWhiteSpace(reply.Text))
                            SetAssistantText(currentStatus, reply.Text);
                        else if (reply.HasOperations)
                        {
                            SetStatus(currentStatus, DescribePendingOperations(reply.Operations));
                            reusableToolStatus = currentStatus;
                        }
                        else
                            SetStatus(currentStatus, "正在整理最终结果…");
                    },
                    (call, result, index, count) =>
                    {
                        if (TryBuildTaskPlan(result, out var plan))
                        {
                            if (activeTaskPlan == null)
                            {
                                activeTaskPlan = new ChatMessage(string.Empty, ChatRole.Assistant);
                                _messages.Add(activeTaskPlan);
                            }
                            ApplyTaskPlan(activeTaskPlan, plan);
                            ScrollToBottom();
                            return;
                        }
                        var target = reusableToolStatus;
                        if (target == null || index > 0)
                            target = new ChatMessage(string.Empty, ChatRole.Assistant);
                        if (!_messages.Contains(target)) _messages.Add(target);
                        ApplyToolResult(target, call, result);
                        ScrollToBottom();
                    });

                chatRounds = run.Rounds;
                chatToolCalls = run.ToolCallCount;
                chatOutcome = run.Completed ? "completed" : "guard_stopped";

                if (!run.Completed)
                {
                    preserveTaskSelectionLock = true;
                    var reason = run.StopReason == "tool_call_limit"
                        ? $"本次已执行 {run.ToolCallCount} 个工具调用"
                        : $"本次已连续运行 {run.Rounds} 轮";
                    AddAssistant($"为避免模型反复调用工具，已暂停自动执行（{reason}）。你可以回复“继续”，我会基于现有结果接着处理。");
                }
            }
            catch (Exception ex)
            {
                chatOutcome = "exception";
                if (thinking.Kind == ChatMessageKind.Status)
                    SetAssistantText(thinking, "分析失败：" + ex.Message);
                else
                    AddAssistant("分析失败：" + ex.Message);
            }
            finally
            {
                chatTimer.Stop();
                PerformanceLogger.Log(
                    "chat_run",
                    chatTimer.ElapsedMilliseconds,
                    "outcome=" + chatOutcome + "|rounds=" + chatRounds +
                    "|tool_calls=" + chatToolCalls + "|stream_chars=" + streamedCharacters +
                    "|stream_renders=" + streamedRenderCount);
                if (!preserveTaskSelectionLock && _app.Selection.LockOwner == "task")
                    _app.Selection.Unlock("task");
                PersistActiveConversation();
                SetInputEnabled(true);
                InputBox.Focus();
            }
        }

        /// <summary>操作执行前的确认回调。返回 true 表示用户同意执行。</summary>
        private bool OnConfirmOperation(string description)
        {
            // 在 UI 线程上弹出确认框;阶段 2 起可换成面板内联确认卡片
            if (!Dispatcher.CheckAccess())
                return (bool)Dispatcher.Invoke(new Func<bool>(() => OnConfirmOperation(description)));

            var ok = MessageBox.Show(
                $"Agent 想执行以下操作,是否允许?\n\n{description}",
                "操作确认",
                MessageBoxButton.OKCancel,
                MessageBoxImage.Question);
            return ok == MessageBoxResult.OK;
        }

        // ---------- 消息渲染辅助 ----------
        private ChatMessage AddUser(string text)
        {
            var m = new ChatMessage(text, ChatRole.User);
            _messages.Add(m);
            ScrollToBottom();
            return m;
        }

        private ChatMessage AddAssistant(string text)
        {
            var m = new ChatMessage(text, ChatRole.Assistant);
            _messages.Add(m);
            ScrollToBottom();
            return m;
        }

        private void CopyMessage_Click(object sender, RoutedEventArgs e)
        {
            var text = ((sender as FrameworkElement)?.DataContext as ChatMessage)?.Text;
            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.BeginInvoke(new Action(() => CopyMessageText(text)));
                return;
            }

            CopyMessageText(text);
        }

        private void CopyMessageText(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                MessageBox.Show("此消息没有可复制的文本。", "复制消息",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            try
            {
                Clipboard.SetText(text);
                MessageBox.Show("消息已复制到剪贴板。", "复制消息",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                // 仅记录异常类型，避免把消息文本或剪贴板内容写入日志。
                ThisAddIn.Log("ChatView: 复制消息失败 (" + ex.GetType().Name + ")");
                MessageBox.Show("复制失败，请重试。", "复制消息",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private ChatMessage AddStatus(string title)
        {
            var message = new ChatMessage(string.Empty, ChatRole.Assistant)
            {
                Kind = ChatMessageKind.Status,
                Title = title
            };
            _messages.Add(message);
            ScrollToBottom();
            return message;
        }

        private static void SetStatus(ChatMessage message, string title)
        {
            message.Kind = ChatMessageKind.Status;
            message.Title = title;
            message.Subtitle = null;
            message.Table = null;
            message.TaskPlan = null;
        }

        private void SetAssistantText(ChatMessage message, string text)
        {
            message.Kind = ChatMessageKind.Text;
            message.Text = text ?? string.Empty;
            message.Title = null;
            message.Subtitle = null;
            message.Table = null;
            message.TaskPlan = null;
            ScrollToBottom();
        }

        private static string DescribePendingOperations(IReadOnlyList<OperationCall> calls)
        {
            if (calls == null || calls.Count == 0) return "正在处理…";
            if (calls.Count > 1) return $"正在执行 {calls.Count} 项操作…";
            switch (calls[0].ToolName)
            {
                case "agent_self_check": return "正在检查 Excel 运行环境…";
                case "cell_read_range": return "正在读取表格数据…";
                case "task_plan": return "正在建立任务计划…";
                case "task_step_update": return "正在更新任务进度…";
                case "data_profile": return "正在体检字段和数据粒度…";
                case "cell_write_range": return "正在写入单元格…";
                case "cell_fill_formula": return "正在填充公式…";
                case "cell_format_range": return "正在设置格式…";
                case "analysis_create_view": return "正在创建安全分析视图…";
                case "chart_create": return "正在创建图表…";
                case "pivot_create": return "正在创建数据透视表…";
                case "dashboard_create": return "正在创建联动数据看板…";
                case "pq_list_queries": return "正在读取 Power Query 查询…";
                case "pq_create_from_range": return "正在创建 Power Query 清洗流程…";
                case "pq_load_to_sheet": return "正在加载 Power Query 结果…";
                case "pq_refresh": return "正在刷新 Power Query…";
                case "pp_list_model": return "正在读取 Power Pivot 数据模型…";
                case "pp_add_query_to_model": return "正在载入 Power Pivot 数据模型…";
                case "pp_refresh_model": return "正在刷新 Power Pivot 数据模型…";
                case "pp_add_relationship": return "正在建立模型关系…";
                case "pp_add_measure": return "正在创建 DAX 度量值…";
                case "pp_create_model_pivot": return "正在创建模型透视表…";
                case "vba_preview_safe": return "正在检查受控 VBA…";
                case "vba_execute_safe": return "正在备份并执行受控 VBA…";
                default: return "正在执行操作…";
            }
        }

        private static void ApplyToolResult(ChatMessage message, OperationCall call, string result)
        {
            string selfCheckSummary;
            if (TryBuildSelfCheckSummary(result, out selfCheckSummary))
            {
                message.Kind = ChatMessageKind.ToolResult;
                message.Title = "Agent 环境自检";
                message.Subtitle = selfCheckSummary;
                message.Text = string.Empty;
                message.Table = null;
                return;
            }
            string profileSummary;
            if (TryBuildProfileSummary(result, out profileSummary))
            {
                message.Kind = ChatMessageKind.ToolResult;
                message.Title = "数据体检完成";
                message.Subtitle = profileSummary;
                message.Text = string.Empty;
                message.Table = null;
                return;
            }
            string modelSummary;
            if (TryBuildModelSummary(result, out modelSummary))
            {
                message.Kind = ChatMessageKind.ToolResult;
                message.Title = "数据模型概览";
                message.Subtitle = modelSummary;
                message.Text = string.Empty;
                message.Table = null;
                return;
            }
            string vbaSummary;
            if (TryBuildVbaPreviewSummary(result, out vbaSummary))
            {
                message.Kind = ChatMessageKind.ToolResult;
                message.Title = "受控 VBA 预览";
                message.Subtitle = vbaSummary;
                message.Text = string.Empty;
                message.Table = null;
                return;
            }
            if (TryBuildTablePreview(result, out var preview))
            {
                message.Kind = ChatMessageKind.TablePreview;
                message.Table = preview;
                message.Text = string.Empty;
                message.Title = null;
                message.Subtitle = null;
                return;
            }

            message.Kind = ChatMessageKind.ToolResult;
            message.Title = GetToolResultTitle(call?.ToolName, result);
            message.Subtitle = result;
            message.Text = string.Empty;
            message.Table = null;
            message.TaskPlan = null;
        }

        private static string GetToolResultTitle(string toolName, string result)
        {
            if (!string.IsNullOrWhiteSpace(result) &&
                (result.Contains("失败") || result.Contains("已跳过") || result.Contains("未确认")))
                return "操作未完成";
            switch (toolName)
            {
                case "agent_self_check": return "环境自检完成";
                case "cell_read_range": return "已读取数据";
                case "task_plan": return "已建立任务计划";
                case "task_step_update": return "已更新任务进度";
                case "data_profile": return "数据体检完成";
                case "cell_write_range": return "已写入单元格";
                case "cell_fill_formula": return "已填充公式";
                case "cell_format_range": return "已应用格式";
                case "analysis_create_view": return "已创建分析视图";
                case "chart_create": return "已创建图表";
                case "pivot_create": return "已创建数据透视表";
                case "dashboard_create": return "已创建联动数据看板";
                case "pq_list_queries": return "已读取 Power Query 查询";
                case "pq_create_from_range": return "已创建 Power Query 查询";
                case "pq_load_to_sheet": return "已加载 Power Query 结果";
                case "pq_refresh": return "已刷新 Power Query";
                case "pp_list_model": return "已读取 Power Pivot 数据模型";
                case "pp_add_query_to_model": return "已载入 Power Pivot 数据模型";
                case "pp_refresh_model": return "已刷新 Power Pivot 数据模型";
                case "pp_add_relationship": return "已建立模型关系";
                case "pp_add_measure": return "已创建 DAX 度量值";
                case "pp_create_model_pivot": return "已创建模型透视表";
                case "vba_preview_safe": return "受控 VBA 已就绪";
                case "vba_execute_safe": return "已执行受控 VBA";
                default: return "操作已完成";
            }
        }

        private static void ApplyTaskPlan(ChatMessage message, TaskPlanData plan)
        {
            message.Kind = ChatMessageKind.TaskPlan;
            message.TaskPlan = plan;
            message.Text = string.Empty;
            message.Title = null;
            message.Subtitle = null;
            message.Table = null;
        }

        private static bool TryBuildTaskPlan(string result, out TaskPlanData plan)
        {
            plan = null;
            if (string.IsNullOrWhiteSpace(result) ||
                !result.StartsWith(TaskExecutionRegistry.PayloadPrefix, StringComparison.Ordinal))
                return false;

            using (var document = JsonDocument.Parse(result.Substring(TaskExecutionRegistry.PayloadPrefix.Length)))
            {
                var root = document.RootElement;
                var steps = new List<TaskPlanStepData>();
                if (root.TryGetProperty("Steps", out var stepArray))
                {
                    foreach (var item in stepArray.EnumerateArray())
                    {
                        steps.Add(new TaskPlanStepData
                        {
                            Index = item.GetProperty("Index").GetInt32(),
                            Title = item.GetProperty("Title").GetString(),
                            Status = item.GetProperty("Status").GetString(),
                            Detail = item.TryGetProperty("Detail", out var detail) && detail.ValueKind == JsonValueKind.String
                                ? detail.GetString()
                                : null
                        });
                    }
                }
                var criteria = new List<string>();
                if (root.TryGetProperty("SuccessCriteria", out var criteriaArray))
                    foreach (var item in criteriaArray.EnumerateArray()) criteria.Add(item.GetString() ?? string.Empty);
                var completed = root.TryGetProperty("CompletedSteps", out var completedValue)
                    ? completedValue.GetInt32() : 0;
                var total = root.TryGetProperty("TotalSteps", out var totalValue)
                    ? totalValue.GetInt32() : steps.Count;
                var isComplete = root.TryGetProperty("IsComplete", out var completeValue) && completeValue.GetBoolean();
                plan = new TaskPlanData
                {
                    Title = root.TryGetProperty("Title", out var titleValue) ? titleValue.GetString() : "执行计划",
                    ProgressText = isComplete ? "全部步骤已经完成并通过完成检查" : $"已完成 {completed}/{total} 步",
                    Badge = isComplete ? "已完成" : $"{completed}/{total}",
                    Steps = steps,
                    SuccessCriteria = criteria
                };
                return true;
            }
        }

        private static bool TryBuildTablePreview(string result, out TablePreviewData preview)
        {
            preview = null;
            if (string.IsNullOrWhiteSpace(result) ||
                !result.StartsWith(ReadRangeOp.TablePayloadPrefix, StringComparison.Ordinal))
                return false;

            using (var document = JsonDocument.Parse(result.Substring(ReadRangeOp.TablePayloadPrefix.Length)))
            {
                var root = document.RootElement;
                var headers = new List<string>();
                if (root.TryGetProperty("headers", out var headerArray))
                    foreach (var header in headerArray.EnumerateArray()) headers.Add(header.GetString() ?? string.Empty);

                var rows = new List<TablePreviewRow>();
                var startRow = root.GetProperty("start_row").GetInt32();
                var rowIndex = 0;
                if (root.TryGetProperty("rows", out var rowArray))
                {
                    foreach (var row in rowArray.EnumerateArray())
                    {
                        // 对话区只展示前 8 行；完整的 12 行结果仍会回传给模型。
                        if (rowIndex >= 8) break;
                        var cells = new List<string>();
                        foreach (var cell in row.EnumerateArray()) cells.Add(cell.GetString() ?? string.Empty);
                        rows.Add(new TablePreviewRow
                        {
                            Number = (startRow + rowIndex).ToString(),
                            Cells = cells,
                            IsAlternate = rowIndex % 2 == 1
                        });
                        rowIndex++;
                    }
                }

                var sheet = root.GetProperty("sheet").GetString();
                var address = root.GetProperty("address").GetString();
                var totalRows = root.GetProperty("total_rows").GetInt32();
                var totalColumns = root.GetProperty("total_columns").GetInt32();
                var shownRows = root.GetProperty("shown_rows").GetInt32();
                var shownColumns = root.GetProperty("shown_columns").GetInt32();
                var truncated = root.GetProperty("truncated").GetBoolean();

                preview = new TablePreviewData
                {
                    Title = "已读取数据",
                    Subtitle = $"{sheet} · {address} · 已读取 {shownRows}/{totalRows} 行、{shownColumns}/{totalColumns} 列",
                    Badge = truncated ? "部分预览" : "完整数据",
                    Headers = headers,
                    Rows = rows
                };
                return true;
            }
        }

        private static bool TryBuildSelfCheckSummary(string result, out string summary)
        {
            summary = null;
            const string prefix = "__AGENT_SELF_CHECK__";
            if (string.IsNullOrWhiteSpace(result) || !result.StartsWith(prefix, StringComparison.Ordinal)) return false;
            try
            {
                using (var document = JsonDocument.Parse(result.Substring(prefix.Length)))
                {
                    var root = document.RootElement;
                    var version = root.TryGetProperty("excel_version", out var versionValue) ? versionValue.GetString() : "未知";
                    var workbookOpen = root.TryGetProperty("workbook_open", out var workbookValue) && workbookValue.GetBoolean();
                    var powerQuery = root.TryGetProperty("power_query", out var queryValue) && queryValue.GetBoolean();
                    var powerPivot = root.TryGetProperty("power_pivot", out var modelValue) && modelValue.GetBoolean();
                    var vba = root.TryGetProperty("vba_project_access", out var vbaValue) && vbaValue.GetBoolean();
                    var warnings = root.TryGetProperty("warnings", out var warningValue) ? warningValue.GetArrayLength() : 0;
                    summary = "Excel " + version + " · " + (workbookOpen ? "工作簿已打开" : "未打开工作簿") +
                              " · Power Query " + (powerQuery ? "可用" : "不可用") +
                              " · Power Pivot " + (powerPivot ? "可用" : "不可用") +
                              " · VBA " + (vba ? "可执行" : "需授权") +
                              (warnings > 0 ? " · " + warnings + " 项提示" : string.Empty);
                    return true;
                }
            }
            catch { return false; }
        }

        private static bool TryBuildVbaPreviewSummary(string result, out string summary)
        {
            summary = null;
            const string prefix = "__AGENT_VBA_PREVIEW__";
            if (string.IsNullOrWhiteSpace(result) || !result.StartsWith(prefix, StringComparison.Ordinal)) return false;
            try
            {
                using (var document = JsonDocument.Parse(result.Substring(prefix.Length)))
                {
                    var root = document.RootElement;
                    summary = root.TryGetProperty("summary", out var value) ? value.GetString() : "白名单宏已通过检查";
                    return true;
                }
            }
            catch { return false; }
        }

        private static bool TryBuildModelSummary(string result, out string summary)
        {
            summary = null;
            const string prefix = "__AGENT_MODEL_LIST__";
            if (string.IsNullOrWhiteSpace(result) || !result.StartsWith(prefix, StringComparison.Ordinal)) return false;
            try
            {
                using (var document = JsonDocument.Parse(result.Substring(prefix.Length)))
                {
                    var root = document.RootElement;
                    var tableCount = root.TryGetProperty("tables", out var tables) ? tables.GetArrayLength() : 0;
                    var relationshipCount = root.TryGetProperty("relationships", out var relationships) ? relationships.GetArrayLength() : 0;
                    var measureCount = root.TryGetProperty("measures", out var measures) ? measures.GetArrayLength() : 0;
                    summary = tableCount + " 张模型表 · " + relationshipCount + " 条关系 · " + measureCount + " 个 DAX 度量值";
                    return true;
                }
            }
            catch { return false; }
        }

        private static bool TryBuildProfileSummary(string result, out string summary)
        {
            summary = null;
            if (string.IsNullOrWhiteSpace(result) ||
                !result.StartsWith(ProfileDataOp.PayloadPrefix, StringComparison.Ordinal))
                return false;

            using (var document = JsonDocument.Parse(result.Substring(ProfileDataOp.PayloadPrefix.Length)))
            {
                var root = document.RootElement;
                var dimensions = 0;
                var measures = 0;
                var times = 0;
                if (root.TryGetProperty("fields", out var fields))
                {
                    foreach (var field in fields.EnumerateArray())
                    {
                        var role = field.TryGetProperty("Role", out var roleValue) ? roleValue.GetString() : string.Empty;
                        if (role == "dimension") dimensions++;
                        else if (role == "measure") measures++;
                        else if (role == "time") times++;
                    }
                }
                var warningCount = root.TryGetProperty("warnings", out var warnings) ? warnings.GetArrayLength() : 0;
                var duplicateRows = root.GetProperty("duplicate_rows").GetInt32();
                summary = $"{root.GetProperty("data_rows").GetInt32()} 行 × {root.GetProperty("columns").GetInt32()} 列 · " +
                          $"时间字段 {times} · 分类维度 {dimensions} · 数值指标 {measures}\n" +
                          $"发现 {warningCount} 项需要关注的问题，其中完全重复记录 {duplicateRows} 条。后续制图会先聚合重复横轴并控制分类数量。";
                return true;
            }
        }

        private void ScrollToBottom()
        {
            if (_scrollToBottomPending) return;
            _scrollToBottomPending = true;
            Dispatcher.BeginInvoke(new Action(() =>
            {
                try { MessagesScroll.ScrollToBottom(); }
                finally { _scrollToBottomPending = false; }
            }), DispatcherPriority.Background);
        }

        private void SetInputEnabled(bool enabled)
        {
            _isBusy = !enabled;
            InputBox.IsEnabled = enabled;
            SendButton.IsEnabled = enabled;
            NewChatButton.IsEnabled = enabled;
            QuickModelCombo.IsEnabled = enabled;
            SelectionReferenceButton.IsEnabled = enabled;
            SelectionLockButton.IsEnabled = enabled;
        }

        private void InputBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (InputPlaceholder != null)
                InputPlaceholder.Visibility = string.IsNullOrEmpty(InputBox.Text)
                    ? Visibility.Visible
                    : Visibility.Collapsed;
        }

        private void SettingsButton_Click(object sender, RoutedEventArgs e)
        {
            SettingsPopup.IsOpen = !SettingsPopup.IsOpen;
        }

        private void ModelSettings_Click(object sender, RoutedEventArgs e)
        {
            OpenSettings("models");
        }

        private void SafetySettings_Click(object sender, RoutedEventArgs e)
        {
            OpenSettings("safety");
        }

        private void WorkbookSettings_Click(object sender, RoutedEventArgs e)
        {
            OpenSettings("workbook");
        }

        private void DiagnosticsSettings_Click(object sender, RoutedEventArgs e)
        {
            OpenSettings("diagnostics");
        }

        private void OpenSettings(string initialSection)
        {
            if (_app == null) return;
            SettingsPopup.IsOpen = false;

            var win = new SettingsWindow(_app.Settings, initialSection);
            var owner = Window.GetWindow(this);
            if (owner != null)
            {
                win.Owner = owner;
            }
            else
            {
                // ElementHost 中通常没有 WPF 顶级窗口，绑定 Excel HWND，
                // 防止设置窗口出现在 Excel 背后。
                var excelWindow = _app.Excel.ActiveWindow;
                if (excelWindow != null)
                    new WindowInteropHelper(win).Owner = new IntPtr(excelWindow.Hwnd);
            }

            if (win.ShowDialog() == true)
            {
                (_app.LLM as OpenAICompatibleClient)?.ReloadConfig(_app.Settings);
                RefreshModelPicker();
                if (_activeConversation != null)
                    _activeConversation.ModelProfileId = _app.Settings.ActiveProfileId;
                PersistActiveConversation();
                UpdateActiveModelLabel();
                UpdateSelectionContextUi();
            }
        }

        private void NewChatButton_Click(object sender, RoutedEventArgs e)
        {
            if (_isBusy || _chatHistory == null) return;
            if (_app.Selection?.LockOwner == "task") _app.Selection.Unlock("task");
            var conversation = _chatStore.CreateConversation(_chatHistory, _app.Settings.ActiveProfileId);
            ActivateConversation(conversation);
            PersistActiveConversation();
        }

        private void QuickModelCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_updatingModelPicker || _app == null || QuickModelCombo.SelectedItem == null) return;
            var profile = (ModelProfile)QuickModelCombo.SelectedItem;
            if (!_app.Settings.SwitchActiveProfile(profile.Id)) return;
            _app.Settings.Save();
            (_app.LLM as OpenAICompatibleClient)?.ReloadConfig(_app.Settings);
            if (_activeConversation != null) _activeConversation.ModelProfileId = profile.Id;
            UpdateActiveModelLabel();
            PersistActiveConversation();
        }

        private void RefreshModelPicker()
        {
            if (_app?.Settings == null || QuickModelCombo == null) return;
            _updatingModelPicker = true;
            QuickModelCombo.ItemsSource = null;
            QuickModelCombo.ItemsSource = _app.Settings.Profiles;
            QuickModelCombo.SelectedItem = _app.Settings.ActiveProfile;
            _updatingModelPicker = false;
        }

        private void ActivateConversation(ChatConversation conversation, bool persistCurrent = true)
        {
            if (conversation == null) return;
            if (persistCurrent) PersistActiveConversation();

            _activeConversation = conversation;
            _chatHistory.ActiveConversationId = conversation.Id;
            _history.Clear();
            if (conversation.History != null) _history.AddRange(conversation.History);
            _messages.Clear();
            if (conversation.Messages != null)
                foreach (var message in conversation.Messages)
                    if (message != null) _messages.Add(message.ToMessage());

            if (!string.IsNullOrWhiteSpace(conversation.ModelProfileId) &&
                _app.Settings.SwitchActiveProfile(conversation.ModelProfileId))
            {
                _app.Settings.Save();
                (_app.LLM as OpenAICompatibleClient)?.ReloadConfig(_app.Settings);
            }
            else
            {
                conversation.ModelProfileId = _app.Settings.ActiveProfileId;
            }

            WelcomePanel.Visibility = _messages.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
            RefreshModelPicker();
            UpdateActiveModelLabel();
            Dispatcher.BeginInvoke(new Action(() =>
            {
                if (_messages.Count == 0) MessagesScroll.ScrollToTop();
                else MessagesScroll.ScrollToBottom();
            }));
        }

        private void PersistActiveConversation()
        {
            if (_chatStore == null || _chatHistory == null || _activeConversation == null) return;
            if (_app?.Settings?.SaveChatHistory == false) return;
            try
            {
                _activeConversation.History = _history.Select(CloneTurn).ToList();
                _activeConversation.Messages = _messages
                    .Select(PersistedChatMessage.FromMessage)
                    .Where(message => message != null)
                    .ToList();
                _activeConversation.ModelProfileId = _app?.Settings?.ActiveProfileId;
                _activeConversation.UpdatedAtUtc = DateTime.UtcNow;
                _chatHistory.ActiveConversationId = _activeConversation.Id;
                _chatStore.Save(_chatHistory);
            }
            catch (Exception ex)
            {
                ThisAddIn.Log("保存对话记录失败: " + ex.Message);
            }
        }

        private static ChatTurn CloneTurn(ChatTurn turn)
        {
            if (turn == null) return null;
            return new ChatTurn
            {
                Role = turn.Role,
                Content = turn.Content,
                ToolCallId = turn.ToolCallId,
                ToolCalls = turn.ToolCalls?.Select(call => new OperationCall
                {
                    CallId = call.CallId,
                    ToolName = call.ToolName,
                    ArgumentsJson = call.ArgumentsJson
                }).ToList()
            };
        }

        private void UpdateConversationTitle(string userText)
        {
            if (_activeConversation == null || _activeConversation.Title != "新对话") return;
            var title = string.Join(" ", (userText ?? string.Empty)
                .Split(new[] { '\r', '\n', '\t' }, StringSplitOptions.RemoveEmptyEntries)).Trim();
            if (title.Length > 24) title = title.Substring(0, 24) + "…";
            _activeConversation.Title = string.IsNullOrWhiteSpace(title) ? "新对话" : title;
        }

        private void ConversationHistory_Click(object sender, RoutedEventArgs e)
        {
            SettingsPopup.IsOpen = false;
            ShowDetails("对话记录");

            var newButton = new Button
            {
                Margin = new Thickness(0, 0, 0, 11),
                HorizontalContentAlignment = HorizontalAlignment.Stretch,
                Content = CreateCardContent("新建对话", "创建独立上下文，当前对话会自动保存。", "新建", "#17734A")
            };
            newButton.SetResourceReference(StyleProperty, "TemplateButtonStyle");
            newButton.Click += NewConversationFromHistory_Click;
            DetailsBody.Children.Add(newButton);

            foreach (var conversation in _chatHistory.Conversations.OrderByDescending(item => item.UpdatedAtUtc))
                DetailsBody.Children.Add(CreateConversationRow(conversation));
        }

        private Grid CreateConversationRow(ChatConversation conversation)
        {
            var row = new Grid { Margin = new Thickness(0, 0, 0, 8) };
            row.ColumnDefinitions.Add(new ColumnDefinition());
            row.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(36) });

            var open = new Button
            {
                Tag = conversation.Id,
                Margin = new Thickness(0, 0, 6, 0),
                HorizontalContentAlignment = HorizontalAlignment.Stretch,
                Content = CreateCardContent(
                    conversation.Title,
                    conversation.UpdatedAtUtc.ToLocalTime().ToString("MM-dd HH:mm") + " · " + GetProfileDisplayName(conversation.ModelProfileId),
                    conversation.Id == _activeConversation?.Id ? "当前" : "打开",
                    conversation.Id == _activeConversation?.Id ? "#17734A" : "#6C756F")
            };
            open.SetResourceReference(StyleProperty, "TemplateButtonStyle");
            open.Click += SwitchConversation_Click;
            row.Children.Add(open);

            var delete = new Button
            {
                Tag = conversation.Id,
                Width = 34,
                Height = 34,
                ToolTip = "删除对话",
                Content = new TextBlock { FontFamily = new FontFamily("Segoe MDL2 Assets"), Text = "\uE74D", FontSize = 12 }
            };
            delete.SetResourceReference(StyleProperty, "IconButtonStyle");
            delete.Click += DeleteConversation_Click;
            Grid.SetColumn(delete, 1);
            row.Children.Add(delete);
            return row;
        }

        private string GetProfileDisplayName(string profileId)
        {
            return _app.Settings.Profiles.FirstOrDefault(profile => profile.Id == profileId)?.DisplayName ?? "默认模型";
        }

        private void NewConversationFromHistory_Click(object sender, RoutedEventArgs e)
        {
            DetailsOverlay.Visibility = Visibility.Collapsed;
            NewChatButton_Click(sender, e);
        }

        private void SwitchConversation_Click(object sender, RoutedEventArgs e)
        {
            var id = (sender as Button)?.Tag as string;
            var conversation = _chatHistory.Conversations.FirstOrDefault(item => item.Id == id);
            if (conversation == null) return;
            DetailsOverlay.Visibility = Visibility.Collapsed;
            ActivateConversation(conversation);
            PersistActiveConversation();
        }

        private void DeleteConversation_Click(object sender, RoutedEventArgs e)
        {
            var id = (sender as Button)?.Tag as string;
            var conversation = _chatHistory.Conversations.FirstOrDefault(item => item.Id == id);
            if (conversation == null) return;
            if (MessageBox.Show("删除对话“" + conversation.Title + "”？此操作不可撤销。", "删除对话",
                    MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes) return;

            var wasActive = conversation.Id == _activeConversation?.Id;
            _chatStore.DeleteConversation(_chatHistory, conversation.Id);
            if (_chatHistory.Conversations.Count == 0)
                _chatStore.CreateConversation(_chatHistory, _app.Settings.ActiveProfileId);
            if (wasActive)
                ActivateConversation(_chatHistory.Conversations.First(), false);
            PersistActiveConversation();
            ConversationHistory_Click(sender, e);
        }

        private void TemplateButton_Click(object sender, RoutedEventArgs e)
        {
            var button = sender as Button;
            ApplySuggestedPrompt(button?.Tag as string);
        }

        private void ComposerPlus_Click(object sender, RoutedEventArgs e)
        {
            ShowCapabilityCenter();
        }

        private void AnalysisTemplates_Click(object sender, RoutedEventArgs e)
        {
            SettingsPopup.IsOpen = false;
            ShowCapabilityCenter();
        }

        private void AnalysisMethods_Click(object sender, RoutedEventArgs e)
        {
            ShowCapabilityCenter();
        }

        private void ShowCapabilityCenter()
        {
            ShowDetails("功能中心");
            string currentCategory = null;
            foreach (var capability in CapabilityCatalog.Items)
            {
                if (!string.Equals(currentCategory, capability.Category, StringComparison.Ordinal))
                {
                    currentCategory = capability.Category;
                    AddCapabilityCategory(currentCategory);
                }
                var available = EditionPolicy.IsCapabilityAvailable(capability.Id, ProductEditionInfo.Current);
                if (available)
                {
                    AddPromptCard(
                        capability.Title,
                        capability.Description,
                        capability.Prompt,
                        capability.Badge,
                        capability.Accent);
                }
                else
                {
                    AddStatusCard(
                        capability.Title,
                        capability.Description,
                        ProductEditionInfo.DisplayName(capability.MinimumEdition),
                        "#8A6A2F");
                }
            }
        }

        private void AddCapabilityCategory(string title)
        {
            DetailsBody.Children.Add(new TextBlock
            {
                Text = title,
                FontSize = 12,
                FontWeight = FontWeights.SemiBold,
                Foreground = (Brush)new BrushConverter().ConvertFrom("#42534A"),
                Margin = new Thickness(2, DetailsBody.Children.Count == 0 ? 0 : 8, 0, 8)
            });
        }

        private void ToolsPermissions_Click(object sender, RoutedEventArgs e)
        {
            SettingsPopup.IsOpen = false;
            ShowDetails("工具与权限");
            AddStatusCard("读取单元格", "可以读取当前工作簿中的指定区域，用于分析上下文。", "已开启", "#17734A");
            AddStatusCard("写入单元格", "可以批量写入普通值；单次操作最多 50,000 个单元格。", "已开启", "#17734A");
            AddStatusCard("公式与格式", "可以填充 A1/R1C1 公式，并设置数字格式、字体、颜色、边框和自适应尺寸。", "已开启", "#17734A");
            AddStatusCard("报告级图表", "可以创建柱形图、条形图、折线图、环形图、饼图、面积图和散点图，并应用统一的报告级样式。", "已开启", "#17734A");
            AddStatusCard("数据体检与制图规则", "制图前识别字段类型、缺失、异常和基数；自动聚合重复横轴、限制分类数量并抑制密集标签。", "已开启", "#17734A");
            AddStatusCard("普通数据透视表", "可以配置行、列、筛选和值字段，并选择求和、计数、平均、最大或最小聚合。", "已开启", "#17734A");
            AddStatusCard("联动数据看板", "默认使用原生切片器和动态透视图；不兼容时可自动回退到下拉筛选，联动由加载项执行且不注入 VBA。源数据保持不变。", "已开启", "#17734A");
            AddStatusCard("Power Query", "可以从当前区域创建可刷新的清洗查询，完成空行、文本、重复项、字段类型、重命名和选列处理，并加载到新工作表。", "MVP 已开启", "#17734A");
            AddStatusCard("Power Pivot / DAX", "可以把 Power Query 载入数据模型，建立表关系、创建 DAX 度量值，并生成跨表模型透视表。", "MVP 已开启", "#17734A");
            AddStatusCard("受控 VBA", "支持刷新全部、自动列宽和导出当前工作表 PDF 三种白名单配方；执行前预览并确认，自动创建备份、移除临时模块并记录审计。", "MVP 已开启", "#17734A");
            AddStatusCard("执行前确认", "当前所有写值、公式和格式操作，默认都会在真正修改工作簿前向你确认。", "已开启", "#17734A");
        }

        private void Appearance_Click(object sender, RoutedEventArgs e)
        {
            SettingsPopup.IsOpen = false;
            ShowDetails("外观");
            AddStatusCard("温和浅色", "当前使用暖白背景、雾绿色状态和低对比边框，适合长时间分析。", "正在使用", "#17734A");
            AddStatusCard("深色主题", "将根据实际使用反馈决定是否加入。", "规划中", "#6C756F");
        }

        private void Help_Click(object sender, RoutedEventArgs e)
        {
            SettingsPopup.IsOpen = false;
            ShowDetails("帮助");
            AddStatusCard("如何开始", "在 Excel 中选中或打开目标工作表，再描述你想检查的数据范围和问题。", "步骤 1", "#17734A");
            AddStatusCard("当前能力", "Agent 可以读写单元格、填充公式、设置格式，并创建报告级图表、普通数据透视表和联动数据看板。", "可用", "#17734A");
            AddStatusCard("安全边界", "所有写操作都需要确认；VBA 只执行预览过的白名单配方，不接受任意代码注入。", "受控写入", "#17734A");
        }

        private void ShowDetails(string title)
        {
            DetailsTitle.Text = title;
            DetailsBody.Children.Clear();
            DetailsOverlay.Visibility = Visibility.Visible;
        }

        private void AddPromptCard(string title, string description, string prompt)
        {
            AddPromptCard(title, description, prompt, "使用", "#17734A");
        }

        private void AddPromptCard(string title, string description, string prompt, string badge, string accent)
        {
            var button = new Button
            {
                Tag = prompt,
                Margin = new Thickness(0, 0, 0, 9),
                HorizontalContentAlignment = HorizontalAlignment.Stretch,
                Content = CreateCardContent(title, description, badge, accent)
            };
            button.SetResourceReference(StyleProperty, "TemplateButtonStyle");
            button.Click += DetailPrompt_Click;
            DetailsBody.Children.Add(button);
        }

        private void AddStatusCard(string title, string description, string status, string statusColor)
        {
            DetailsBody.Children.Add(new Border
            {
                Background = (Brush)new BrushConverter().ConvertFrom("#FFFEFC"),
                BorderBrush = (Brush)new BrushConverter().ConvertFrom("#DFE5E1"),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(11),
                Padding = new Thickness(12),
                Margin = new Thickness(0, 0, 0, 9),
                Child = CreateCardContent(title, description, status, statusColor)
            });
        }

        private static Grid CreateCardContent(string title, string description, string status, string statusColor)
        {
            var grid = new Grid();
            grid.ColumnDefinitions.Add(new ColumnDefinition());
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            var copy = new StackPanel();
            copy.Children.Add(new TextBlock
            {
                Text = title,
                FontSize = 13,
                FontWeight = FontWeights.SemiBold,
                Foreground = (Brush)new BrushConverter().ConvertFrom("#17271F")
            });
            copy.Children.Add(new TextBlock
            {
                Text = description,
                FontSize = 10.5,
                Foreground = (Brush)new BrushConverter().ConvertFrom("#64716A"),
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 4, 10, 0),
                LineHeight = 16
            });
            grid.Children.Add(copy);

            var statusBlock = new TextBlock
            {
                Text = status,
                FontSize = 10.5,
                Foreground = (Brush)new BrushConverter().ConvertFrom(statusColor),
                VerticalAlignment = VerticalAlignment.Top,
                Margin = new Thickness(8, 1, 0, 0)
            };
            Grid.SetColumn(statusBlock, 1);
            grid.Children.Add(statusBlock);
            return grid;
        }

        private void DetailPrompt_Click(object sender, RoutedEventArgs e)
        {
            ApplySuggestedPrompt((sender as Button)?.Tag as string);
            DetailsOverlay.Visibility = Visibility.Collapsed;
        }

        private void ApplySuggestedPrompt(string prompt)
        {
            if (string.IsNullOrWhiteSpace(prompt)) return;
            InputBox.Text = prompt;
            InputBox.CaretIndex = InputBox.Text.Length;
            InputBox.Focus();
        }

        private void CloseDetails_Click(object sender, RoutedEventArgs e)
        {
            DetailsOverlay.Visibility = Visibility.Collapsed;
        }

        private void UpdateActiveModelLabel()
        {
            if (_app?.Settings == null) return;
            ActiveModelText.Text = _app.Settings.ActiveProfile.DisplayName + " · " + _app.Settings.Model;
        }

        private void SelectionContext_Changed(object sender, EventArgs e)
        {
            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.BeginInvoke(new Action(UpdateSelectionContextUi));
                return;
            }
            UpdateSelectionContextUi();
        }

        private void UpdateSelectionContextUi()
        {
            if (_app?.Selection == null || SelectionContextText == null) return;
            var selection = _app.Selection.Effective;
            SelectionContextText.Text = selection?.DisplayText ?? "未检测到 Excel 选区";
            WorkbookContextText.Text = selection?.IsValid == true ? selection.WorkbookName : "当前工作簿";

            var locked = _app.Selection.IsLocked;
            SelectionLockIcon.Text = locked ? "\uE785" : "\uE72E";
            SelectionLockIcon.Foreground = (Brush)new BrushConverter().ConvertFrom(locked ? "#168653" : "#62736A");
            SelectionLockButton.Background = (Brush)new BrushConverter().ConvertFrom(locked ? "#E5F3EA" : "#F2F7F4");
            SelectionLockButton.ToolTip = locked
                ? (_app.Selection.LockOwner == "task" ? "任务已锁定该选区，点击可解除" : "选区已锁定，点击可解除")
                : "锁定当前选区";
            SelectionReferenceButton.ToolTip = locked
                ? "点击引用已锁定选区"
                : "点击在输入框中引用当前选区";

            switch (_app.Settings.AutomationMode)
            {
                case "auto": AutomationModeText.Text = "自动执行"; break;
                case "ask_every_time": AutomationModeText.Text = "每次询问"; break;
                case "custom": AutomationModeText.Text = "自定义权限"; break;
                default: AutomationModeText.Text = "安全自动"; break;
            }
        }

        private void SelectionReference_Click(object sender, RoutedEventArgs e)
        {
            if (_app?.Selection?.Effective?.IsValid != true)
            {
                MessageBox.Show("请先在 Excel 中选中要引用的区域。", "当前选区",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            const string token = "@当前选区";
            var position = Math.Max(0, InputBox.CaretIndex);
            var prefix = position > 0 && !char.IsWhiteSpace(InputBox.Text[position - 1]) ? " " : string.Empty;
            var suffix = position < InputBox.Text.Length && !char.IsWhiteSpace(InputBox.Text[position]) ? " " : string.Empty;
            var insertion = prefix + token + suffix;
            InputBox.Text = InputBox.Text.Insert(position, insertion);
            InputBox.CaretIndex = position + insertion.Length;
            InputBox.Focus();
        }

        private void SelectionLock_Click(object sender, RoutedEventArgs e)
        {
            if (_app?.Selection == null || _isBusy) return;
            if (_app.Selection.IsLocked)
                _app.Selection.Unlock();
            else if (!_app.Selection.LockCurrent("manual"))
                MessageBox.Show("请先在 Excel 中选中要锁定的区域。", "锁定选区",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            UpdateSelectionContextUi();
        }
    }
}
