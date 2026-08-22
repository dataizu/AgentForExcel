using System;
using Microsoft.Office.Interop.Excel;

namespace AgentForExcel.Operations
{
    /// <summary>
    /// 批量执行作用域:进入时关闭屏幕刷新 / 事件 / 自动重算,退出时恢复原值。
    /// 用于会多次写单元格的重操作(画像素、建看板、格式化、建分析页),
    /// 避免每次赋值都触发重绘与事件链,耗时成倍放大。
    /// 恢复必须用进入时保存的原值而非硬编码默认,避免吞掉用户的手动计算设置。
    ///
    /// 三个属性各自独立记录"是否改动过"、独立恢复:构造中途失败(如受保护视图下
    /// EnableEvents 赋值抛出)不能把 ScreenUpdating 永久留在 false;Dispose 时
    /// 任一属性恢复失败也不影响其余恢复 —— 否则 Excel 会表现为"假死"不重绘,
    /// 或用户的手动计算设置被悄悄改成 Manual。
    /// </summary>
    public sealed class ExcelBatchScope : IDisposable
    {
        private readonly Application _excel;
        private readonly bool _screenUpdating;
        private readonly bool _enableEvents;
        private readonly XlCalculation _calculation;
        private bool _screenUpdatingChanged;
        private bool _enableEventsChanged;
        private bool _calculationChanged;

        public ExcelBatchScope(AppContext context)
        {
            if (context == null) throw new ArgumentNullException(nameof(context));
            _excel = context.Excel;
            try
            {
                _screenUpdating = _excel.ScreenUpdating;
                _excel.ScreenUpdating = false;
                _screenUpdatingChanged = true;
            }
            catch { }
            try
            {
                _enableEvents = _excel.EnableEvents;
                _excel.EnableEvents = false;
                _enableEventsChanged = true;
            }
            catch { }
            try
            {
                _calculation = _excel.Calculation;
                if (_calculation != XlCalculation.xlCalculationManual)
                {
                    _excel.Calculation = XlCalculation.xlCalculationManual;
                    _calculationChanged = true;
                }
            }
            catch { }
        }

        public void Dispose()
        {
            if (_screenUpdatingChanged)
            {
                try { _excel.ScreenUpdating = _screenUpdating; } catch { }
            }
            if (_enableEventsChanged)
            {
                try { _excel.EnableEvents = _enableEvents; } catch { }
            }
            if (_calculationChanged)
            {
                try
                {
                    _excel.Calculation = _calculation;
                    // 切回自动计算会触发一次异步重算;同步算完再离开作用域,
                    // 避免调用方的下一个 COM 调用撞上 Excel 忙(0x800AC472)。
                    if (_calculation == XlCalculation.xlCalculationAutomatic)
                        _excel.Calculate();
                }
                catch { }
            }
        }
    }
}
