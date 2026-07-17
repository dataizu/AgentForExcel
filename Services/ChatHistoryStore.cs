using System;
using System.IO;
using System.Linq;
using System.Text.Json;
using AgentForExcel.Models;

namespace AgentForExcel.Services
{
    /// <summary>把多会话聊天记录持久化到本机用户目录。</summary>
    public sealed class ChatHistoryStore
    {
        private const int MaxConversations = 50;
        private readonly string _path;

        public ChatHistoryStore(string path = null)
        {
            _path = string.IsNullOrWhiteSpace(path) ? DefaultPath : path;
        }

        public string StoragePath => _path;

        public ChatHistoryDocument Load()
        {
            if (!File.Exists(_path)) return new ChatHistoryDocument();
            try
            {
                var document = JsonSerializer.Deserialize<ChatHistoryDocument>(File.ReadAllText(_path), JsonOptions)
                               ?? new ChatHistoryDocument();
                document.Conversations = document.Conversations ?? new System.Collections.Generic.List<ChatConversation>();
                foreach (var conversation in document.Conversations)
                {
                    conversation.History = conversation.History ?? new System.Collections.Generic.List<AI.ChatTurn>();
                    conversation.Messages = conversation.Messages ?? new System.Collections.Generic.List<PersistedChatMessage>();
                    if (string.IsNullOrWhiteSpace(conversation.Title)) conversation.Title = "新对话";
                }
                if (!document.Conversations.Any(item => item.Id == document.ActiveConversationId))
                    document.ActiveConversationId = document.Conversations.FirstOrDefault()?.Id;
                return document;
            }
            catch
            {
                // 保留损坏文件供人工恢复；当前会话从空记录开始。
                return new ChatHistoryDocument();
            }
        }

        public ChatConversation CreateConversation(ChatHistoryDocument document, string modelProfileId)
        {
            if (document == null) throw new ArgumentNullException(nameof(document));
            var conversation = new ChatConversation { ModelProfileId = modelProfileId };
            document.Conversations.Insert(0, conversation);
            document.ActiveConversationId = conversation.Id;
            return conversation;
        }

        public bool DeleteConversation(ChatHistoryDocument document, string conversationId)
        {
            if (document == null || string.IsNullOrWhiteSpace(conversationId)) return false;
            var conversation = document.Conversations.FirstOrDefault(item => item.Id == conversationId);
            if (conversation == null) return false;
            document.Conversations.Remove(conversation);
            if (document.ActiveConversationId == conversationId)
                document.ActiveConversationId = document.Conversations.FirstOrDefault()?.Id;
            return true;
        }

        public void Save(ChatHistoryDocument document)
        {
            if (document == null) throw new ArgumentNullException(nameof(document));
            document.Version = 1;
            document.Conversations = document.Conversations
                .OrderByDescending(item => item.UpdatedAtUtc)
                .Take(MaxConversations)
                .ToList();
            var directory = Path.GetDirectoryName(_path);
            if (!string.IsNullOrWhiteSpace(directory)) Directory.CreateDirectory(directory);
            var temporaryPath = _path + ".tmp";
            File.WriteAllText(temporaryPath, JsonSerializer.Serialize(document, JsonOptions));
            if (File.Exists(_path)) File.Replace(temporaryPath, _path, null);
            else File.Move(temporaryPath, _path);
        }

        private static string DefaultPath => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "AgentForExcel", "conversations.json");

        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            PropertyNameCaseInsensitive = true,
            WriteIndented = true
        };
    }
}
