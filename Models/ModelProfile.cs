using System;

namespace AgentForExcel.Models
{
    /// <summary>一个可独立切换的大模型连接档案。</summary>
    public sealed class ModelProfile : NotificationObject
    {
        private string _id = Guid.NewGuid().ToString("N");
        private string _displayName = "新模型";
        private string _providerName = "DeepSeek";
        private string _apiKey = string.Empty;
        private string _baseUrl = "https://api.deepseek.com";
        private string _model = "deepseek-v4-flash";
        private double _temperature = 0.3;

        public string Id { get => _id; set => SetField(ref _id, string.IsNullOrWhiteSpace(value) ? Guid.NewGuid().ToString("N") : value); }
        public string DisplayName { get => _displayName; set => SetField(ref _displayName, string.IsNullOrWhiteSpace(value) ? "未命名模型" : value.Trim()); }
        public string ProviderName
        {
            get => _providerName;
            set { if (SetField(ref _providerName, value ?? string.Empty)) OnPropertyChanged(nameof(Summary)); }
        }
        public string ApiKey { get => _apiKey; set => SetField(ref _apiKey, value ?? string.Empty); }
        public string BaseUrl { get => _baseUrl; set => SetField(ref _baseUrl, value ?? string.Empty); }
        public string Model
        {
            get => _model;
            set { if (SetField(ref _model, value ?? string.Empty)) OnPropertyChanged(nameof(Summary)); }
        }
        public double Temperature { get => _temperature; set => SetField(ref _temperature, Math.Max(0, Math.Min(1, value))); }

        public string Summary => ProviderName + " · " + Model;

        public ModelProfile Clone(string id = null, string displayName = null)
        {
            return new ModelProfile
            {
                Id = string.IsNullOrWhiteSpace(id) ? Id : id,
                DisplayName = string.IsNullOrWhiteSpace(displayName) ? DisplayName : displayName,
                ProviderName = ProviderName,
                ApiKey = ApiKey,
                BaseUrl = BaseUrl,
                Model = Model,
                Temperature = Temperature
            };
        }
    }
}
