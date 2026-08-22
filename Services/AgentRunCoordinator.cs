using System;

namespace AgentForExcel.Services
{
    /// <summary>
    /// 进程级 Agent 运行互斥。
    ///
    /// 架构上每个 Excel 窗口各有一个对话面板,但运行时共享状态
    /// (任务选区锁、静态任务计划、聊天历史文件)都是单例:
    /// 两个窗口并行发消息会互相覆盖选区锁、清掉对方的计划、
    /// 并用各自的会话快照把对方的新会话从文件里抹掉。
    /// 因此同一 Excel 进程同一时刻只允许一个对话运行 ——
    /// 这也符合 Excel COM 单线程执行的现实。
    ///
    /// 所有调用都发生在 Excel 主线程,普通字段即可,无需加锁。
    /// </summary>
    public sealed class AgentRunCoordinator
    {
        private object _owner;
        private string _description;

        /// <summary>是否有对话正在运行。</summary>
        public bool IsRunning => _owner != null;

        /// <summary>
        /// 尝试获取运行权。返回 false 表示另一个窗口正在运行,
        /// runningDescription 携带那边的消息开头,供提示文案使用。
        /// </summary>
        public bool TryBegin(object owner, string description, out string runningDescription)
        {
            if (_owner != null && !ReferenceEquals(_owner, owner))
            {
                runningDescription = _description;
                return false;
            }
            _owner = owner;
            _description = description;
            runningDescription = null;
            return true;
        }

        /// <summary>释放运行权。只有当前持有者能释放,防止误清他人的运行。</summary>
        public void End(object owner)
        {
            if (ReferenceEquals(_owner, owner))
            {
                _owner = null;
                _description = null;
            }
        }
    }
}
