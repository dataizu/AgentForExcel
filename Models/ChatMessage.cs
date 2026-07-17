using System.Collections.Generic;
using System.Windows;
using System.Windows.Media;
using System.Windows.Controls;

namespace AgentForExcel.Models
{
    /// <summary>消息角色。</summary>
    public enum ChatRole { User, Assistant }

    public enum ChatMessageKind { Text, Status, ToolResult, TablePreview, TaskPlan }

    /// <summary>
    /// 一条对话消息,供 WPF 绑定。
    /// 气泡外观改用 Role 判断 + DataTemplate 里的触发器实现,
    /// 不再依赖 Application.Current.TryFindResource(在 VSTO 宿主中为 null)。
    /// </summary>
    public class ChatMessage : NotificationObject
    {
        private string _text;
        private ChatMessageKind _kind;
        private string _title;
        private string _subtitle;
        private TablePreviewData _table;
        private TaskPlanData _taskPlan;

        public ChatMessage(string text, ChatRole role)
        {
            Text = text;
            Role = role;
            Kind = ChatMessageKind.Text;
        }

        public ChatRole Role { get; }

        public string Text
        {
            get => _text;
            set { _text = value; OnPropertyChanged(); }
        }

        public ChatMessageKind Kind
        {
            get => _kind;
            set => SetField(ref _kind, value);
        }

        public string Title
        {
            get => _title;
            set => SetField(ref _title, value);
        }

        public string Subtitle
        {
            get => _subtitle;
            set => SetField(ref _subtitle, value);
        }

        public TablePreviewData Table
        {
            get => _table;
            set => SetField(ref _table, value);
        }

        public TaskPlanData TaskPlan
        {
            get => _taskPlan;
            set => SetField(ref _taskPlan, value);
        }

        /// <summary>是否为用户消息(DataTemplate 触发器据此切换气泡外观)。</summary>
        public bool IsUser => Role == ChatRole.User;

        /// <summary>用户消息的前景色(白色)。</summary>
        public Brush Foreground => Role == ChatRole.User
            ? Brushes.White
            : (Brush)new BrushConverter().ConvertFrom("#1F2937");
    }

    public sealed class TablePreviewData
    {
        public string Title { get; set; }
        public string Subtitle { get; set; }
        public string Badge { get; set; }
        public IReadOnlyList<string> Headers { get; set; }
        public IReadOnlyList<TablePreviewRow> Rows { get; set; }
    }

    public sealed class TaskPlanData
    {
        public string Title { get; set; }
        public string ProgressText { get; set; }
        public string Badge { get; set; }
        public IReadOnlyList<TaskPlanStepData> Steps { get; set; }
        public IReadOnlyList<string> SuccessCriteria { get; set; }
    }

    public sealed class TaskPlanStepData
    {
        public int Index { get; set; }
        public string Title { get; set; }
        public string Status { get; set; }
        public string Detail { get; set; }

        public string Glyph
        {
            get
            {
                switch (Status)
                {
                    case "completed": return "✓";
                    case "failed": return "!";
                    case "in_progress": return "•";
                    default: return "○";
                }
            }
        }

        public Brush StatusBrush
        {
            get
            {
                var color = Status == "completed" ? "#17734A" :
                    Status == "failed" ? "#C2413A" :
                    Status == "in_progress" ? "#D18A2E" : "#89938E";
                return (Brush)new BrushConverter().ConvertFrom(color);
            }
        }
    }

    public sealed class TablePreviewRow
    {
        public string Number { get; set; }
        public IReadOnlyList<string> Cells { get; set; }
        public bool IsAlternate { get; set; }
    }
}
