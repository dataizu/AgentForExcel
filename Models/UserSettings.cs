using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text.Json;

namespace AgentForExcel.Models
{
    /// <summary>用户配置：支持多个模型档案和运行时快速切换。</summary>
    public class UserSettings : NotificationObject
    {
        private const string DefaultBaseUrl = "https://api.deepseek.com";
        private const string DefaultModel = "deepseek-v4-flash";
        private const string DefaultProviderName = "DeepSeek";

        private readonly ObservableCollection<ModelProfile> _profiles = new ObservableCollection<ModelProfile>();
        private string _activeProfileId;
        private bool _requireConfirmOnWrite = true;
        private string _automationMode = "safe_auto";
        private bool _autoAllowNewSheetOutputs = true;
        private bool _autoAllowSelectedBlankWrites = true;
        private bool _autoAllowFormattingInSelection = true;
        private int _autoWriteMaxCells = 5000;
        private bool _preserveSourceData = true;
        private bool _preferNewWorksheetForOutputs = true;
        private bool _enablePowerQuery = true;
        private bool _enablePowerPivot = true;
        private bool _enableVba = true;
        private bool _saveChatHistory = true;
        private string _defaultAnalysisScope = "CurrentRegion";

        public UserSettings()
        {
            EnsureProfile();
        }

        public ObservableCollection<ModelProfile> Profiles => _profiles;

        public string ActiveProfileId
        {
            get => _activeProfileId;
            private set
            {
                if (!SetField(ref _activeProfileId, value)) return;
                RaiseActiveProfileProperties();
            }
        }

        public ModelProfile ActiveProfile
        {
            get
            {
                EnsureProfile();
                return _profiles.FirstOrDefault(profile =>
                           string.Equals(profile.Id, _activeProfileId, StringComparison.OrdinalIgnoreCase))
                       ?? _profiles[0];
            }
        }

        // 兼容现有 AI 客户端与调用方：这些属性始终代理到当前档案。
        public string ApiKey { get => ActiveProfile.ApiKey; set { ActiveProfile.ApiKey = value; OnPropertyChanged(); } }
        public string BaseUrl { get => ActiveProfile.BaseUrl; set { ActiveProfile.BaseUrl = value; OnPropertyChanged(); } }
        public string Model { get => ActiveProfile.Model; set { ActiveProfile.Model = value; OnPropertyChanged(); } }
        public string ProviderName { get => ActiveProfile.ProviderName; set { ActiveProfile.ProviderName = value; OnPropertyChanged(); } }
        public double Temperature { get => ActiveProfile.Temperature; set { ActiveProfile.Temperature = value; OnPropertyChanged(); } }

        public bool RequireConfirmOnWrite
        {
            get => _requireConfirmOnWrite;
            set => SetField(ref _requireConfirmOnWrite, value);
        }

        /// <summary>ask_every_time / safe_auto / custom。</summary>
        public string AutomationMode
        {
            get => _automationMode;
            set
            {
                var normalized = value == "ask_every_time" || value == "custom" ? value : "safe_auto";
                SetField(ref _automationMode, normalized);
            }
        }

        public bool AutoAllowNewSheetOutputs
        {
            get => _autoAllowNewSheetOutputs;
            set => SetField(ref _autoAllowNewSheetOutputs, value);
        }

        public bool AutoAllowSelectedBlankWrites
        {
            get => _autoAllowSelectedBlankWrites;
            set => SetField(ref _autoAllowSelectedBlankWrites, value);
        }

        public bool AutoAllowFormattingInSelection
        {
            get => _autoAllowFormattingInSelection;
            set => SetField(ref _autoAllowFormattingInSelection, value);
        }

        public int AutoWriteMaxCells
        {
            get => _autoWriteMaxCells;
            set => SetField(ref _autoWriteMaxCells, Math.Max(1, Math.Min(50000, value)));
        }

        public bool PreserveSourceData
        {
            get => _preserveSourceData;
            set => SetField(ref _preserveSourceData, value);
        }

        public bool PreferNewWorksheetForOutputs
        {
            get => _preferNewWorksheetForOutputs;
            set => SetField(ref _preferNewWorksheetForOutputs, value);
        }

        public bool EnablePowerQuery
        {
            get => _enablePowerQuery;
            set => SetField(ref _enablePowerQuery, value);
        }

        public bool EnablePowerPivot
        {
            get => _enablePowerPivot;
            set => SetField(ref _enablePowerPivot, value);
        }

        public bool EnableVba
        {
            get => _enableVba;
            set => SetField(ref _enableVba, value);
        }

        public bool SaveChatHistory
        {
            get => _saveChatHistory;
            set => SetField(ref _saveChatHistory, value);
        }

        public string DefaultAnalysisScope
        {
            get => _defaultAnalysisScope;
            set
            {
                var normalized = value == "Selection" || value == "UsedRange" ? value : "CurrentRegion";
                SetField(ref _defaultAnalysisScope, normalized);
            }
        }

        public bool SwitchActiveProfile(string profileId)
        {
            var profile = _profiles.FirstOrDefault(item =>
                string.Equals(item.Id, profileId, StringComparison.OrdinalIgnoreCase));
            if (profile == null) return false;
            ActiveProfileId = profile.Id;
            return true;
        }

        public void ReplaceProfiles(IEnumerable<ModelProfile> profiles, string activeProfileId)
        {
            _profiles.Clear();
            if (profiles != null)
            {
                foreach (var profile in profiles)
                    if (profile != null) _profiles.Add(profile.Clone());
            }
            EnsureProfile();
            ActiveProfileId = _profiles.Any(profile =>
                string.Equals(profile.Id, activeProfileId, StringComparison.OrdinalIgnoreCase))
                ? activeProfileId
                : _profiles[0].Id;
            RaiseActiveProfileProperties();
        }

        public static UserSettings Load() => LoadFrom(StorePath);

        public static UserSettings LoadFrom(string path)
        {
            var settings = new UserSettings();
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path)) return settings;
            try
            {
                var json = File.ReadAllText(path);
                var document = JsonSerializer.Deserialize<SettingsDocument>(json, JsonOptions);
                if (document?.Profiles != null && document.Profiles.Count > 0)
                {
                    settings.ReplaceProfiles(document.Profiles.Select(ToProfile), document.ActiveProfileId);
                    settings.RequireConfirmOnWrite = document.RequireConfirmOnWrite;
                    settings.AutomationMode = document.AutomationMode;
                    settings.AutoAllowNewSheetOutputs = document.AutoAllowNewSheetOutputs;
                    settings.AutoAllowSelectedBlankWrites = document.AutoAllowSelectedBlankWrites;
                    settings.AutoAllowFormattingInSelection = document.AutoAllowFormattingInSelection;
                    settings.AutoWriteMaxCells = document.AutoWriteMaxCells;
                    settings.PreserveSourceData = document.PreserveSourceData;
                    settings.PreferNewWorksheetForOutputs = document.PreferNewWorksheetForOutputs;
                    settings.EnablePowerQuery = document.EnablePowerQuery;
                    settings.EnablePowerPivot = document.EnablePowerPivot;
                    settings.EnableVba = document.EnableVba;
                    settings.SaveChatHistory = document.SaveChatHistory;
                    settings.DefaultAnalysisScope = document.DefaultAnalysisScope;
                    return settings;
                }

                // 自动迁移旧版单模型 settings.json。
                using (var legacy = JsonDocument.Parse(json))
                {
                    var root = legacy.RootElement;
                    var profile = settings.ActiveProfile;
                    profile.ApiKey = ReadLegacyString(root, "ApiKey", profile.ApiKey);
                    profile.BaseUrl = ReadLegacyString(root, "BaseUrl", profile.BaseUrl);
                    profile.Model = ReadLegacyString(root, "Model", profile.Model);
                    profile.ProviderName = ReadLegacyString(root, "ProviderName", profile.ProviderName);
                    profile.DisplayName = profile.ProviderName + " · " + profile.Model;
                    double temperature;
                    if (double.TryParse(ReadLegacyString(root, "Temperature", "0.3"), out temperature))
                        profile.Temperature = temperature;
                    bool requireConfirm;
                    if (bool.TryParse(ReadLegacyString(root, "RequireConfirmOnWrite", "True"), out requireConfirm))
                        settings.RequireConfirmOnWrite = requireConfirm;
                }
            }
            catch
            {
                // 首次运行或旧文件损坏时保留安全默认值。
            }
            return settings;
        }

        public void Save() => SaveTo(StorePath);

        public void SaveTo(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("配置路径不能为空。", nameof(path));
            EnsureProfile();
            var document = new SettingsDocument
            {
                Version = 4,
                ActiveProfileId = ActiveProfile.Id,
                RequireConfirmOnWrite = RequireConfirmOnWrite,
                AutomationMode = AutomationMode,
                AutoAllowNewSheetOutputs = AutoAllowNewSheetOutputs,
                AutoAllowSelectedBlankWrites = AutoAllowSelectedBlankWrites,
                AutoAllowFormattingInSelection = AutoAllowFormattingInSelection,
                AutoWriteMaxCells = AutoWriteMaxCells,
                PreserveSourceData = PreserveSourceData,
                PreferNewWorksheetForOutputs = PreferNewWorksheetForOutputs,
                EnablePowerQuery = EnablePowerQuery,
                EnablePowerPivot = EnablePowerPivot,
                EnableVba = EnableVba,
                SaveChatHistory = SaveChatHistory,
                DefaultAnalysisScope = DefaultAnalysisScope,
                Profiles = _profiles.Select(profile => new ModelProfileDocument
                {
                    Id = profile.Id,
                    DisplayName = profile.DisplayName,
                    ProviderName = profile.ProviderName,
                    ApiKey = profile.ApiKey,
                    BaseUrl = profile.BaseUrl,
                    Model = profile.Model,
                    Temperature = profile.Temperature
                }).ToList()
            };
            WriteAtomically(path, JsonSerializer.Serialize(document, JsonOptions));
        }

        private void EnsureProfile()
        {
            if (_profiles.Count == 0)
            {
                _profiles.Add(new ModelProfile
                {
                    DisplayName = "DeepSeek 默认",
                    ProviderName = DefaultProviderName,
                    BaseUrl = DefaultBaseUrl,
                    Model = DefaultModel,
                    Temperature = 0.3
                });
            }
            if (string.IsNullOrWhiteSpace(_activeProfileId) ||
                !_profiles.Any(profile => string.Equals(profile.Id, _activeProfileId, StringComparison.OrdinalIgnoreCase)))
                _activeProfileId = _profiles[0].Id;
        }

        private void RaiseActiveProfileProperties()
        {
            OnPropertyChanged(nameof(ActiveProfile));
            OnPropertyChanged(nameof(ApiKey));
            OnPropertyChanged(nameof(BaseUrl));
            OnPropertyChanged(nameof(Model));
            OnPropertyChanged(nameof(ProviderName));
            OnPropertyChanged(nameof(Temperature));
        }

        private static ModelProfile ToProfile(ModelProfileDocument profile)
        {
            return new ModelProfile
            {
                Id = profile.Id,
                DisplayName = profile.DisplayName,
                ProviderName = profile.ProviderName,
                ApiKey = profile.ApiKey,
                BaseUrl = profile.BaseUrl,
                Model = profile.Model,
                Temperature = profile.Temperature
            };
        }

        private static string ReadLegacyString(JsonElement root, string name, string fallback)
        {
            if (!root.TryGetProperty(name, out var value)) return fallback;
            return value.ValueKind == JsonValueKind.String ? value.GetString() : value.ToString();
        }

        private static void WriteAtomically(string path, string content)
        {
            var directory = Path.GetDirectoryName(path);
            if (!string.IsNullOrWhiteSpace(directory)) Directory.CreateDirectory(directory);
            var temporaryPath = path + ".tmp";
            File.WriteAllText(temporaryPath, content);
            if (File.Exists(path)) File.Replace(temporaryPath, path, null);
            else File.Move(temporaryPath, path);
        }

        private static string StorePath => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "AgentForExcel", "settings.json");

        private sealed class SettingsDocument
        {
            public int Version { get; set; }
            public string ActiveProfileId { get; set; }
            public bool RequireConfirmOnWrite { get; set; } = true;
            public string AutomationMode { get; set; } = "safe_auto";
            public bool AutoAllowNewSheetOutputs { get; set; } = true;
            public bool AutoAllowSelectedBlankWrites { get; set; } = true;
            public bool AutoAllowFormattingInSelection { get; set; } = true;
            public int AutoWriteMaxCells { get; set; } = 5000;
            public bool PreserveSourceData { get; set; } = true;
            public bool PreferNewWorksheetForOutputs { get; set; } = true;
            public bool EnablePowerQuery { get; set; } = true;
            public bool EnablePowerPivot { get; set; } = true;
            public bool EnableVba { get; set; } = true;
            public bool SaveChatHistory { get; set; } = true;
            public string DefaultAnalysisScope { get; set; } = "CurrentRegion";
            public List<ModelProfileDocument> Profiles { get; set; }
        }

        private sealed class ModelProfileDocument
        {
            public string Id { get; set; }
            public string DisplayName { get; set; }
            public string ProviderName { get; set; }
            public string ApiKey { get; set; }
            public string BaseUrl { get; set; }
            public string Model { get; set; }
            public double Temperature { get; set; }
        }

        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            PropertyNameCaseInsensitive = true,
            WriteIndented = true
        };
    }
}
