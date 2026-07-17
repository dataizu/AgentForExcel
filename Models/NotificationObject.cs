using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace AgentForExcel.Models
{
    /// <summary>
    /// INotifyPropertyChanged 基类。所有需要 WPF 数据绑定的模型继承它。
    /// </summary>
    public abstract class NotificationObject : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged([CallerMemberName] string propertyName = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));

        /// <summary>设置字段值并触发属性变更通知(若值未变则跳过)。</summary>
        protected bool SetField<T>(ref T field, T value, [CallerMemberName] string propertyName = null)
        {
            if (Equals(field, value)) return false;
            field = value;
            OnPropertyChanged(propertyName);
            return true;
        }
    }
}
