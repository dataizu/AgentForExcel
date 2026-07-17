using System;
using AgentForExcel.Models;
using Microsoft.Office.Interop.Excel;

namespace AgentForExcel.Services
{
    /// <summary>跟踪活动选区，并允许在 Agent 任务期间冻结目标范围。</summary>
    public sealed class SelectionContextService
    {
        private readonly Application _excel;
        private SelectionContext _current = new SelectionContext();
        private SelectionContext _locked;

        public SelectionContextService(Application excel)
        {
            _excel = excel ?? throw new ArgumentNullException(nameof(excel));
        }

        public event EventHandler Changed;

        public SelectionContext Current => _current;
        public SelectionContext Locked => _locked;
        public SelectionContext Effective => _locked ?? _current;
        public bool IsLocked => _locked != null;
        public string LockOwner { get; private set; }

        public void Refresh()
        {
            try
            {
                var range = _excel.Selection as Range;
                var sheet = _excel.ActiveSheet as Worksheet;
                Update(sheet, range);
            }
            catch
            {
                Update(null, null);
            }
        }

        public void Update(Worksheet sheet, Range range)
        {
            var next = Capture(sheet, range);
            if (Same(_current, next)) return;
            _current = next;
            Changed?.Invoke(this, EventArgs.Empty);
        }

        public bool LockCurrent(string owner)
        {
            Refresh();
            if (!_current.IsValid) return false;
            _locked = _current.Clone();
            LockOwner = string.IsNullOrWhiteSpace(owner) ? "manual" : owner;
            Changed?.Invoke(this, EventArgs.Empty);
            return true;
        }

        public void Unlock(string owner = null)
        {
            if (_locked == null) return;
            if (!string.IsNullOrWhiteSpace(owner) &&
                !string.Equals(owner, LockOwner, StringComparison.OrdinalIgnoreCase)) return;
            _locked = null;
            LockOwner = null;
            Refresh();
            Changed?.Invoke(this, EventArgs.Empty);
        }

        private static SelectionContext Capture(Worksheet sheet, Range range)
        {
            if (sheet == null || range == null) return new SelectionContext();
            try
            {
                var workbook = sheet.Parent as Workbook;
                string fullName;
                try { fullName = workbook?.FullName; }
                catch { fullName = workbook?.Name; }
                return new SelectionContext
                {
                    WorkbookName = workbook?.Name,
                    WorkbookFullName = fullName,
                    SheetName = sheet.Name,
                    Address = range.Address,
                    RowCount = range.Rows.Count,
                    ColumnCount = range.Columns.Count,
                    CellCount = Convert.ToInt64(range.CountLarge),
                    IsMultiArea = range.Areas.Count > 1,
                    CapturedAtUtc = DateTime.UtcNow
                };
            }
            catch
            {
                return new SelectionContext();
            }
        }

        private static bool Same(SelectionContext left, SelectionContext right)
        {
            return string.Equals(left?.WorkbookFullName, right?.WorkbookFullName, StringComparison.OrdinalIgnoreCase) &&
                   string.Equals(left?.SheetName, right?.SheetName, StringComparison.OrdinalIgnoreCase) &&
                   string.Equals(left?.Address, right?.Address, StringComparison.OrdinalIgnoreCase) &&
                   left?.CellCount == right?.CellCount && left?.IsMultiArea == right?.IsMultiArea;
        }
    }
}
