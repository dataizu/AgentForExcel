using System;
using System.Collections.Generic;
using System.Linq;
using AgentForExcel.AI;

namespace AgentForExcel.Models
{
    public sealed class ChatHistoryDocument
    {
        public int Version { get; set; } = 1;
        public string ActiveConversationId { get; set; }
        public List<ChatConversation> Conversations { get; set; } = new List<ChatConversation>();
    }

    public sealed class ChatConversation
    {
        public string Id { get; set; } = Guid.NewGuid().ToString("N");
        public string Title { get; set; } = "新对话";
        public string ModelProfileId { get; set; }
        public DateTime CreatedAtUtc { get; set; } = DateTime.UtcNow;
        public DateTime UpdatedAtUtc { get; set; } = DateTime.UtcNow;
        public List<ChatTurn> History { get; set; } = new List<ChatTurn>();
        public List<PersistedChatMessage> Messages { get; set; } = new List<PersistedChatMessage>();
    }

    public sealed class PersistedChatMessage
    {
        public ChatRole Role { get; set; }
        public ChatMessageKind Kind { get; set; }
        public string Text { get; set; }
        public string Title { get; set; }
        public string Subtitle { get; set; }
        public PersistedTablePreview Table { get; set; }
        public PersistedTaskPlan TaskPlan { get; set; }

        public static PersistedChatMessage FromMessage(ChatMessage message)
        {
            if (message == null || message.Kind == ChatMessageKind.Status) return null;
            return new PersistedChatMessage
            {
                Role = message.Role,
                Kind = message.Kind,
                Text = message.Text,
                Title = message.Title,
                Subtitle = message.Subtitle,
                Table = PersistedTablePreview.FromPreview(message.Table),
                TaskPlan = PersistedTaskPlan.FromPlan(message.TaskPlan)
            };
        }

        public ChatMessage ToMessage()
        {
            return new ChatMessage(Text ?? string.Empty, Role)
            {
                Kind = Kind,
                Title = Title,
                Subtitle = Subtitle,
                Table = Table?.ToPreview(),
                TaskPlan = TaskPlan?.ToPlan()
            };
        }
    }

    public sealed class PersistedTablePreview
    {
        public string Title { get; set; }
        public string Subtitle { get; set; }
        public string Badge { get; set; }
        public List<string> Headers { get; set; } = new List<string>();
        public List<PersistedTableRow> Rows { get; set; } = new List<PersistedTableRow>();

        public static PersistedTablePreview FromPreview(TablePreviewData preview)
        {
            if (preview == null) return null;
            return new PersistedTablePreview
            {
                Title = preview.Title,
                Subtitle = preview.Subtitle,
                Badge = preview.Badge,
                Headers = preview.Headers?.ToList() ?? new List<string>(),
                Rows = preview.Rows?.Select(row => new PersistedTableRow
                {
                    Number = row.Number,
                    Cells = row.Cells?.ToList() ?? new List<string>(),
                    IsAlternate = row.IsAlternate
                }).ToList() ?? new List<PersistedTableRow>()
            };
        }

        public TablePreviewData ToPreview()
        {
            return new TablePreviewData
            {
                Title = Title,
                Subtitle = Subtitle,
                Badge = Badge,
                Headers = Headers ?? new List<string>(),
                Rows = Rows?.Select(row => new TablePreviewRow
                {
                    Number = row.Number,
                    Cells = row.Cells ?? new List<string>(),
                    IsAlternate = row.IsAlternate
                }).ToList() ?? new List<TablePreviewRow>()
            };
        }
    }

    public sealed class PersistedTableRow
    {
        public string Number { get; set; }
        public List<string> Cells { get; set; } = new List<string>();
        public bool IsAlternate { get; set; }
    }

    public sealed class PersistedTaskPlan
    {
        public string Title { get; set; }
        public string ProgressText { get; set; }
        public string Badge { get; set; }
        public List<PersistedTaskStep> Steps { get; set; } = new List<PersistedTaskStep>();
        public List<string> SuccessCriteria { get; set; } = new List<string>();

        public static PersistedTaskPlan FromPlan(TaskPlanData plan)
        {
            if (plan == null) return null;
            return new PersistedTaskPlan
            {
                Title = plan.Title,
                ProgressText = plan.ProgressText,
                Badge = plan.Badge,
                SuccessCriteria = plan.SuccessCriteria?.ToList() ?? new List<string>(),
                Steps = plan.Steps?.Select(step => new PersistedTaskStep
                {
                    Index = step.Index,
                    Title = step.Title,
                    Status = step.Status,
                    Detail = step.Detail
                }).ToList() ?? new List<PersistedTaskStep>()
            };
        }

        public TaskPlanData ToPlan()
        {
            return new TaskPlanData
            {
                Title = Title,
                ProgressText = ProgressText,
                Badge = Badge,
                SuccessCriteria = SuccessCriteria ?? new List<string>(),
                Steps = Steps?.Select(step => new TaskPlanStepData
                {
                    Index = step.Index,
                    Title = step.Title,
                    Status = step.Status,
                    Detail = step.Detail
                }).ToList() ?? new List<TaskPlanStepData>()
            };
        }
    }

    public sealed class PersistedTaskStep
    {
        public int Index { get; set; }
        public string Title { get; set; }
        public string Status { get; set; }
        public string Detail { get; set; }
    }
}
