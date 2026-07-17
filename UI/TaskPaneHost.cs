using System.Windows.Forms;

namespace AgentForExcel.UI
{
    /// <summary>
    /// 承载 WPF 内容(Elemen"tHost)的 WinForms UserControl。
    /// VSTO 的 CustomTaskPane 要求传入 UserControl,因此这里做一层薄包装。
    /// </summary>
    public partial class TaskPaneHost : UserControl
    {
        public TaskPaneHost()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            this.Name = "TaskPaneHost";
            this.Size = new System.Drawing.Size(360, 600);
            this.BackColor = System.Drawing.Color.White;
            this.ResumeLayout(false);
        }
    }
}
