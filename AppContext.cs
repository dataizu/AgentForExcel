using System;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools;

namespace AgentForExcel
{
    /// <summary>
    /// 全局应用上下文。集中持有 Excel 对象模型与各服务实例,
    /// 操作层 / UI 层统一通过此对象访问资源,避免散落的 Globals 调用。
    /// </summary>
    public sealed class AppContext : IDisposable
    {
        /// <summary>Excel Application 实例。</summary>
        public Application Excel { get; }

        /// <summary>VSTO Factory(创建原生宿主项用)。</summary>
        public Factory Factory { get; }

        /// <summary>AI 客户端(阶段 1 接入)。</summary>
        public AI.ILLMClient LLM { get; private set; }

        /// <summary>操作派发器(阶段 2 起填充能力)。</summary>
        public Operations.OperationDispatcher Dispatcher { get; private set; }

        /// <summary>用户配置(API Key / BaseUrl / 模型名)。</summary>
        public Models.UserSettings Settings { get; private set; }

        /// <summary>实时选区与任务级选区锁定。</summary>
        public Services.SelectionContextService Selection { get; private set; }

        /// <summary>写操作权限与风险分级策略。</summary>
        public Services.PermissionPolicyService Permissions { get; private set; }

        public AppContext(Application excel, Factory factory)
        {
            Excel = excel ?? throw new ArgumentNullException(nameof(excel));
            Factory = factory ?? throw new ArgumentNullException(nameof(factory));
        }

        /// <summary>初始化各服务。从持久化配置读取,构造 AI 客户端与派发器。</summary>
        public void Initialize()
        {
            Settings = Models.UserSettings.Load();

            Selection = new Services.SelectionContextService(Excel);
            Selection.Refresh();
            Permissions = new Services.PermissionPolicyService(this);

            // AI 客户端:OpenAI 兼容协议,可对接智谱 GLM / DeepSeek / 通义等
            LLM = new AI.OpenAICompatibleClient(Settings);

            // 操作派发器:阶段 2 起注册各能力模块
            Dispatcher = new Operations.OperationDispatcher();
            Operations.OperationRegistration.RegisterAll(Dispatcher, this);
        }

        public void Dispose()
        {
            (LLM as IDisposable)?.Dispose();
        }
    }
}
