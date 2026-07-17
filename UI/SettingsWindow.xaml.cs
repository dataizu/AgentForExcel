using System;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using AgentForExcel.Models;
using AgentForExcel.Operations;

namespace AgentForExcel.UI
{
    /// <summary>多模型档案管理窗口。</summary>
    public partial class SettingsWindow : Window
    {
        private readonly UserSettings _settings;
        private readonly ObservableCollection<ModelProfile> _profiles = new ObservableCollection<ModelProfile>();
        private ModelProfile _selectedProfile;
        private bool _initializing;

        private sealed class ProviderPreset
        {
            public string Name { get; }
            public string BaseUrl { get; }
            public string DefaultModel { get; }
            public string[] Models { get; }
            public bool IsCustom { get; }

            public ProviderPreset(string name, string baseUrl, string defaultModel, params string[] models)
            {
                Name = name; BaseUrl = baseUrl; DefaultModel = defaultModel; Models = models ?? new string[0];
            }

            private ProviderPreset()
            {
                Name = "自定义（OpenAI 兼容）"; BaseUrl = ""; DefaultModel = ""; Models = new string[0]; IsCustom = true;
            }

            public static ProviderPreset Custom() => new ProviderPreset();
        }

        private static readonly ProviderPreset[] Presets =
        {
            new ProviderPreset("DeepSeek", "https://api.deepseek.com", "deepseek-v4-flash", "deepseek-v4-flash", "deepseek-v4-pro"),
            new ProviderPreset("MiniMax", "https://api.minimaxi.com/v1", "MiniMax-M2.7", "MiniMax-M2.7", "MiniMax-M2.7-highspeed", "MiniMax-M2.5", "MiniMax-M2.5-highspeed", "MiniMax-M2.1"),
            new ProviderPreset("智谱 GLM（国内）", "https://open.bigmodel.cn/api/paas/v4", "glm-5.2", "glm-5.2", "glm-5.1", "glm-5-turbo", "glm-5", "glm-4.7", "glm-4.7-flash"),
            new ProviderPreset("通义千问 Qwen", "https://dashscope.aliyuncs.com/compatible-mode/v1", "qwen-plus", "qwen-plus", "qwen-max", "qwen-flash"),
            new ProviderPreset("OpenAI / ChatGPT", "https://api.openai.com/v1", "gpt-5.6", "gpt-5.6", "gpt-5.6-terra", "gpt-5.6-luna", "gpt-5.4-mini"),
            new ProviderPreset("Anthropic / Claude", "https://api.anthropic.com/v1", "claude-sonnet-4-6", "claude-opus-4-8", "claude-sonnet-4-6", "claude-haiku-4-5"),
            new ProviderPreset("Google / Gemini", "https://generativelanguage.googleapis.com/v1beta/openai", "gemini-3.5-flash", "gemini-3.5-flash", "gemini-3.1-pro-preview", "gemini-flash-lite-latest"),
            ProviderPreset.Custom()
        };

        public SettingsWindow(UserSettings settings, string initialSection = "models")
        {
            InitializeComponent();
            _settings = settings ?? throw new ArgumentNullException(nameof(settings));
            _initializing = true;
            foreach (var profile in settings.Profiles) _profiles.Add(profile.Clone());
            ProfileList.ItemsSource = _profiles;
            ProviderCombo.ItemsSource = Presets;
            AutomationModeCombo.SelectedValue = settings.AutomationMode;
            AutoNewSheetCheck.IsChecked = settings.AutoAllowNewSheetOutputs;
            AutoSelectedWritesCheck.IsChecked = settings.AutoAllowSelectedBlankWrites;
            AutoFormatCheck.IsChecked = settings.AutoAllowFormattingInSelection;
            AutoWriteLimitBox.Text = settings.AutoWriteMaxCells.ToString();
            PreserveSourceCheck.IsChecked = settings.PreserveSourceData;
            NewSheetCheck.IsChecked = settings.PreferNewWorksheetForOutputs;
            EnablePowerQueryCheck.IsChecked = settings.EnablePowerQuery;
            EnablePowerPivotCheck.IsChecked = settings.EnablePowerPivot;
            EnableVbaCheck.IsChecked = settings.EnableVba;
            SaveHistoryCheck.IsChecked = settings.SaveChatHistory;
            DefaultScopeCombo.SelectedValue = settings.DefaultAnalysisScope;
            StoragePathText.Text = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "AgentForExcel");
            TempSlider.ValueChanged += TempSlider_ValueChanged;

            _selectedProfile = _profiles.FirstOrDefault(profile => profile.Id == settings.ActiveProfileId) ?? _profiles.First();
            ProfileList.SelectedItem = _selectedProfile;
            LoadEditor(_selectedProfile);
            _initializing = false;
            ShowSection(initialSection);
        }

        private void Navigation_Click(object sender, RoutedEventArgs e)
        {
            ShowSection((sender as Button)?.Tag as string);
        }

        private void ShowSection(string section)
        {
            var selected = string.IsNullOrWhiteSpace(section) ? "models" : section;
            ModelsPanel.Visibility = selected == "models" ? Visibility.Visible : Visibility.Collapsed;
            SafetyPanel.Visibility = selected == "safety" ? Visibility.Visible : Visibility.Collapsed;
            WorkbookPanel.Visibility = selected == "workbook" ? Visibility.Visible : Visibility.Collapsed;
            PrivacyPanel.Visibility = selected == "privacy" ? Visibility.Visible : Visibility.Collapsed;
            AppearancePanel.Visibility = selected == "appearance" ? Visibility.Visible : Visibility.Collapsed;
            DiagnosticsPanel.Visibility = selected == "diagnostics" ? Visibility.Visible : Visibility.Collapsed;

            var active = (Brush)new BrushConverter().ConvertFrom("#E7F2E9");
            var normal = Brushes.Transparent;
            var buttons = new[] { ModelsNav, SafetyNav, WorkbookNav, PrivacyNav, AppearanceNav, DiagnosticsNav };
            foreach (var button in buttons)
            {
                button.Background = string.Equals(button.Tag as string, selected, StringComparison.OrdinalIgnoreCase)
                    ? active
                    : normal;
                button.FontWeight = string.Equals(button.Tag as string, selected, StringComparison.OrdinalIgnoreCase)
                    ? FontWeights.SemiBold
                    : FontWeights.Normal;
            }
        }

        private void ProfileList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_initializing || ProfileList.SelectedItem == null) return;
            CommitEditor();
            _selectedProfile = (ModelProfile)ProfileList.SelectedItem;
            LoadEditor(_selectedProfile);
        }

        private void LoadEditor(ModelProfile profile)
        {
            if (profile == null) return;
            _initializing = true;
            DisplayNameBox.Text = profile.DisplayName;
            ApiKeyBox.Password = profile.ApiKey ?? string.Empty;
            BaseUrlBox.Text = profile.BaseUrl ?? string.Empty;
            ModelCombo.Text = profile.Model ?? string.Empty;
            TempSlider.Value = profile.Temperature;
            TempValue.Text = profile.Temperature.ToString("0.0");
            ProviderCombo.SelectedIndex = FindPresetIndex(profile.ProviderName, profile.BaseUrl);
            ApplyPreset((ProviderPreset)ProviderCombo.SelectedItem, profile.BaseUrl, profile.Model);
            _initializing = false;
        }

        private void CommitEditor()
        {
            if (_selectedProfile == null) return;
            var preset = ProviderCombo.SelectedItem as ProviderPreset ?? Presets[Presets.Length - 1];
            _selectedProfile.DisplayName = string.IsNullOrWhiteSpace(DisplayNameBox.Text)
                ? (string.IsNullOrWhiteSpace(ModelCombo.Text) ? "未命名模型" : ModelCombo.Text.Trim())
                : DisplayNameBox.Text.Trim();
            _selectedProfile.ProviderName = preset.Name;
            _selectedProfile.ApiKey = ApiKeyBox.Password?.Trim();
            _selectedProfile.BaseUrl = string.IsNullOrWhiteSpace(BaseUrlBox.Text) ? preset.BaseUrl : BaseUrlBox.Text.Trim();
            _selectedProfile.Model = string.IsNullOrWhiteSpace(ModelCombo.Text) ? preset.DefaultModel : ModelCombo.Text.Trim();
            _selectedProfile.Temperature = TempSlider.Value;
        }

        private void AddProfile_Click(object sender, RoutedEventArgs e)
        {
            CommitEditor();
            var profile = new ModelProfile
            {
                DisplayName = "新模型",
                ProviderName = Presets[0].Name,
                BaseUrl = Presets[0].BaseUrl,
                Model = Presets[0].DefaultModel,
                Temperature = 0.3
            };
            _profiles.Add(profile);
            ProfileList.SelectedItem = profile;
            ProfileList.ScrollIntoView(profile);
        }

        private void DuplicateProfile_Click(object sender, RoutedEventArgs e)
        {
            if (_selectedProfile == null) return;
            CommitEditor();
            var copy = _selectedProfile.Clone(Guid.NewGuid().ToString("N"), _selectedProfile.DisplayName + " 副本");
            _profiles.Add(copy);
            ProfileList.SelectedItem = copy;
            ProfileList.ScrollIntoView(copy);
        }

        private void DeleteProfile_Click(object sender, RoutedEventArgs e)
        {
            if (_selectedProfile == null) return;
            if (_profiles.Count <= 1)
            {
                MessageBox.Show("至少保留一个模型档案。", "模型与连接", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            if (MessageBox.Show("删除模型档案“" + _selectedProfile.DisplayName + "”？", "模型与连接",
                    MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes) return;
            var index = _profiles.IndexOf(_selectedProfile);
            _profiles.Remove(_selectedProfile);
            ProfileList.SelectedIndex = Math.Max(0, Math.Min(index, _profiles.Count - 1));
        }

        private void ProviderCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_initializing || ProviderCombo.SelectedItem == null) return;
            ApplyPreset((ProviderPreset)ProviderCombo.SelectedItem);
        }

        private void ApplyPreset(ProviderPreset preset, string preferredBaseUrl = null, string preferredModel = null)
        {
            if (preset == null) return;
            ModelCombo.ItemsSource = preset.Models;
            BaseUrlBox.Text = preset.IsCustom
                ? preferredBaseUrl ?? BaseUrlBox.Text
                : string.IsNullOrWhiteSpace(preferredBaseUrl) ? preset.BaseUrl : preferredBaseUrl;
            ModelCombo.Text = preset.IsCustom
                ? preferredModel ?? ModelCombo.Text
                : string.IsNullOrWhiteSpace(preferredModel) ? preset.DefaultModel : preferredModel;
        }

        private static int FindPresetIndex(string providerName, string baseUrl)
        {
            for (var i = 0; i < Presets.Length - 1; i++)
                if (string.Equals(Presets[i].Name, providerName, StringComparison.OrdinalIgnoreCase) || SameUrl(Presets[i].BaseUrl, baseUrl))
                    return i;
            return Presets.Length - 1;
        }

        private static bool SameUrl(string left, string right) => string.Equals(
            (left ?? "").Trim().TrimEnd('/'), (right ?? "").Trim().TrimEnd('/'), StringComparison.OrdinalIgnoreCase);

        private void TempSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (TempValue != null) TempValue.Text = e.NewValue.ToString("0.0");
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            CommitEditor();
            var invalid = _profiles.FirstOrDefault(profile => string.IsNullOrWhiteSpace(profile.BaseUrl) || string.IsNullOrWhiteSpace(profile.Model));
            if (invalid != null)
            {
                MessageBox.Show("模型档案“" + invalid.DisplayName + "”缺少 Base URL 或模型 ID。", "模型与连接",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                ProfileList.SelectedItem = invalid;
                return;
            }
            _settings.ReplaceProfiles(_profiles, _selectedProfile?.Id);
            int autoWriteLimit;
            if (!int.TryParse(AutoWriteLimitBox.Text, out autoWriteLimit) || autoWriteLimit < 1 || autoWriteLimit > 50000)
            {
                MessageBox.Show("单次自动写入上限需要是 1 到 50000 之间的整数。", "执行与安全",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                ShowSection("safety");
                AutoWriteLimitBox.Focus();
                return;
            }

            _settings.AutomationMode = AutomationModeCombo.SelectedValue as string;
            _settings.RequireConfirmOnWrite = _settings.AutomationMode == "ask_every_time";
            _settings.AutoAllowNewSheetOutputs = AutoNewSheetCheck.IsChecked ?? true;
            _settings.AutoAllowSelectedBlankWrites = AutoSelectedWritesCheck.IsChecked ?? true;
            _settings.AutoAllowFormattingInSelection = AutoFormatCheck.IsChecked ?? true;
            _settings.AutoWriteMaxCells = autoWriteLimit;
            _settings.PreserveSourceData = PreserveSourceCheck.IsChecked ?? true;
            _settings.PreferNewWorksheetForOutputs = NewSheetCheck.IsChecked ?? true;
            _settings.EnablePowerQuery = EnablePowerQueryCheck.IsChecked ?? true;
            _settings.EnablePowerPivot = EnablePowerPivotCheck.IsChecked ?? true;
            _settings.EnableVba = EnableVbaCheck.IsChecked ?? true;
            _settings.SaveChatHistory = SaveHistoryCheck.IsChecked ?? true;
            _settings.DefaultAnalysisScope = DefaultScopeCombo.SelectedValue as string;
            _settings.Save();
            DialogResult = true;
            Close();
        }

        private async void RunDiagnostics_Click(object sender, RoutedEventArgs e)
        {
            var button = sender as Button;
            if (button != null) button.IsEnabled = false;
            DiagnosticsResult.Text = "正在检查 Excel 环境…";
            try
            {
                var results = await ThisAddIn.App.Dispatcher.ExecuteAsync(
                    new[]
                    {
                        new OperationCall
                        {
                            CallId = "settings-diagnostics",
                            ToolName = "agent_self_check",
                            ArgumentsJson = "{}"
                        }
                    },
                    _ => false);
                DiagnosticsResult.Text = results.FirstOrDefault() ?? "自检完成，但没有返回详细信息。";
            }
            catch (Exception ex)
            {
                DiagnosticsResult.Text = "自检失败：" + ex.Message;
            }
            finally
            {
                if (button != null) button.IsEnabled = true;
            }
        }

        private void OpenLogFolder_Click(object sender, RoutedEventArgs e)
        {
            var directory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "AgentForExcel");
            Directory.CreateDirectory(directory);
            Process.Start(new ProcessStartInfo("explorer.exe", "\"" + directory + "\"")
            {
                UseShellExecute = true
            });
        }

        private void Cancel_Click(object sender, RoutedEventArgs e) => Close();
    }
}
