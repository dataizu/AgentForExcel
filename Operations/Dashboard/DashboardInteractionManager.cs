using System;
using System.Collections.Generic;
using System.Globalization;
using Microsoft.Office.Interop.Excel;

namespace AgentForExcel.Operations.Dashboard
{
    internal sealed class DashboardFilterBinding
    {
        public string DashboardSheet { get; set; }
        public string SupportSheet { get; set; }
        public string ControlAddress { get; set; }
        public string FieldName { get; set; }
        public string ListName { get; set; }
    }

    /// <summary>
    /// 为下拉兼容模式提供工作簿级联动。逻辑驻留在已签名的 VSTO 加载项中，
    /// 不向用户工作簿注入 VBA；绑定定义写入隐藏辅助页，因此重新打开文件后仍可恢复。
    /// </summary>
    internal static class DashboardInteractionManager
    {
        internal const string AllSelection = "（全部）";
        private const string DefinitionPrefix = "AgentDashboardBinding_";
        private const string DefinitionMarker = "__AGENT_DASHBOARD_FILTER_V1__";
        private const int MetadataStartColumn = 56; // BD

        private static Microsoft.Office.Interop.Excel.Application _application;
        private static bool _handlingChange;

        internal static void Initialize(Microsoft.Office.Interop.Excel.Application application)
        {
            if (application == null) return;
            if (ReferenceEquals(_application, application)) return;
            Shutdown();
            _application = application;
            _application.SheetChange += ApplicationSheetChange;
        }

        internal static void Shutdown()
        {
            if (_application != null)
            {
                try { _application.SheetChange -= ApplicationSheetChange; } catch { }
            }
            _application = null;
            _handlingChange = false;
        }

        internal static string RegisterDashboard(
            Microsoft.Office.Interop.Excel.Application application,
            Workbook workbook,
            Worksheet supportSheet,
            string token,
            IList<DashboardFilterBinding> bindings)
        {
            Initialize(application);
            var first = supportSheet.Cells[1, MetadataStartColumn];
            var last = supportSheet.Cells[bindings.Count + 1, MetadataStartColumn + 4];
            var metadata = supportSheet.Range[first, last];
            metadata.ClearContents();
            ((Range)metadata.Cells[1, 1]).Value2 = DefinitionMarker;
            ((Range)metadata.Cells[1, 2]).Value2 = "DashboardSheet";
            ((Range)metadata.Cells[1, 3]).Value2 = "ControlAddress";
            ((Range)metadata.Cells[1, 4]).Value2 = "FieldName";
            ((Range)metadata.Cells[1, 5]).Value2 = "ListName";
            for (var index = 0; index < bindings.Count; index++)
            {
                var binding = bindings[index];
                ((Range)metadata.Cells[index + 2, 1]).Value2 = binding.DashboardSheet;
                ((Range)metadata.Cells[index + 2, 2]).Value2 = binding.SupportSheet;
                ((Range)metadata.Cells[index + 2, 3]).Value2 = binding.ControlAddress;
                ((Range)metadata.Cells[index + 2, 4]).Value2 = binding.FieldName;
                ((Range)metadata.Cells[index + 2, 5]).Value2 = binding.ListName;
            }

            var definitionName = DefinitionPrefix + token;
            workbook.Names.Add(definitionName,
                "='" + supportSheet.Name.Replace("'", "''") + "'!" + metadata.Address, false);
            return definitionName;
        }

        internal static void DeleteDefinition(Workbook workbook, string definitionName)
        {
            if (workbook == null || string.IsNullOrWhiteSpace(definitionName)) return;
            try
            {
                dynamic definition = workbook.Names.Item(definitionName);
                Range metadata = null;
                try { metadata = definition.RefersToRange; } catch { }
                if (metadata != null)
                {
                    for (var row = 2; row <= metadata.Rows.Count; row++)
                    {
                        var listName = Convert.ToString(((Range)metadata.Cells[row, 5]).Value2);
                        if (!string.IsNullOrWhiteSpace(listName))
                            Try(() => workbook.Names.Item(listName).Delete());
                    }
                }
                definition.Delete();
            }
            catch { }
        }

        internal static void RefreshDashboard(Workbook workbook, string dashboardSheetName)
        {
            if (workbook == null || string.IsNullOrWhiteSpace(dashboardSheetName)) return;
            foreach (var definition in ReadDefinitions(workbook))
            {
                if (!string.Equals(definition.DashboardSheet, dashboardSheetName, StringComparison.OrdinalIgnoreCase))
                    continue;
                ApplyDefinition(workbook, definition);
            }
        }

        private static void ApplicationSheetChange(object sheetObject, Range target)
        {
            if (_handlingChange || target == null) return;
            var sheet = sheetObject as Worksheet;
            if (sheet == null) return;
            var workbook = sheet.Parent as Workbook;
            if (workbook == null) return;

            try
            {
                foreach (var definition in ReadDefinitions(workbook))
                {
                    if (!string.Equals(definition.DashboardSheet, sheet.Name, StringComparison.OrdinalIgnoreCase))
                        continue;
                    var controlRange = sheet.Range[definition.ControlAddress];
                    if (_application.Intersect(target, controlRange) == null) continue;
                    _handlingChange = true;
                    ApplyDashboardDefinitions(workbook, definition.DashboardSheet);
                    break;
                }
            }
            catch (Exception ex)
            {
                ThisAddIn.Log("下拉看板联动失败: " + ex.Message);
            }
            finally
            {
                _handlingChange = false;
            }
        }

        private static void ApplyDashboardDefinitions(Workbook workbook, string dashboardSheetName)
        {
            var definitions = new List<DashboardFilterBinding>();
            foreach (var definition in ReadDefinitions(workbook))
                if (string.Equals(definition.DashboardSheet, dashboardSheetName, StringComparison.OrdinalIgnoreCase))
                    definitions.Add(definition);
            if (definitions.Count == 0) return;

            var dashboardSheet = (Worksheet)workbook.Worksheets.Item[dashboardSheetName];
            var supportSheet = (Worksheet)workbook.Worksheets.Item[definitions[0].SupportSheet];
            dynamic dashboardPivots = dashboardSheet.PivotTables();
            for (var index = 1; index <= Convert.ToInt32(dashboardPivots.Count); index++)
                ApplyFilters((PivotTable)dashboardPivots.Item(index), dashboardSheet, definitions);
            dynamic supportPivots = supportSheet.PivotTables();
            for (var index = 1; index <= Convert.ToInt32(supportPivots.Count); index++)
                ApplyFilters((PivotTable)supportPivots.Item(index), dashboardSheet, definitions);
            Try(() => workbook.Application.Calculate());
        }

        private static void ApplyDefinition(Workbook workbook, DashboardFilterBinding definition)
        {
            ApplyDashboardDefinitions(workbook, definition.DashboardSheet);
        }

        private static void ApplyFilters(PivotTable pivot, Worksheet dashboardSheet, IList<DashboardFilterBinding> definitions)
        {
            foreach (var definition in definitions)
            {
                var selection = Convert.ToString(((Range)dashboardSheet.Range[definition.ControlAddress].Cells[1, 1]).Value2)?.Trim();
                ApplyFilter(pivot, definition.FieldName, selection);
            }
            pivot.RefreshTable();
        }

        private static void ApplyFilter(PivotTable pivot, string fieldName, string selection)
        {
            var field = (PivotField)pivot.PivotFields(fieldName);
            var showAll = string.IsNullOrWhiteSpace(selection) || selection == AllSelection;
            Try(() => field.ClearAllFilters());

            if (field.Orientation == XlPivotFieldOrientation.xlHidden)
            {
                field.Orientation = XlPivotFieldOrientation.xlPageField;
                field.EnableMultiplePageItems = false;
            }

            if (field.Orientation == XlPivotFieldOrientation.xlPageField)
            {
                if (!showAll)
                    SetPageSelection(field, selection);
                return;
            }

            dynamic items = field.PivotItems();
            if (showAll)
            {
                for (var index = 1; index <= items.Count; index++)
                {
                    var item = (PivotItem)items.Item(index);
                    Try(() => item.Visible = true);
                }
                return;
            }

            PivotItem selected = null;
            for (var index = 1; index <= items.Count; index++)
            {
                var item = (PivotItem)items.Item(index);
                if (Matches(item, selection))
                {
                    selected = item;
                    break;
                }
            }
            if (selected == null)
                throw new InvalidOperationException("筛选字段「" + fieldName + "」中找不到「" + selection + "」。");

            selected.Visible = true;
            for (var index = 1; index <= items.Count; index++)
            {
                var item = (PivotItem)items.Item(index);
                if (!Matches(item, selection)) Try(() => item.Visible = false);
            }
        }

        private static void SetPageSelection(PivotField field, string selection)
        {
            dynamic items = field.PivotItems();
            for (var index = 1; index <= items.Count; index++)
            {
                var item = (PivotItem)items.Item(index);
                if (!Matches(item, selection)) continue;
                field.CurrentPage = item.Name;
                return;
            }
            throw new InvalidOperationException("筛选字段「" + field.Name + "」中找不到「" + selection + "」。");
        }

        private static bool Matches(PivotItem item, string selection)
        {
            if (string.Equals(Convert.ToString(item.Name), selection, StringComparison.CurrentCultureIgnoreCase)) return true;
            try
            {
                if (string.Equals(Convert.ToString(item.Caption), selection, StringComparison.CurrentCultureIgnoreCase)) return true;
            }
            catch { }
            try
            {
                if (string.Equals(Convert.ToString(item.Value, CultureInfo.InvariantCulture), selection,
                    StringComparison.InvariantCultureIgnoreCase)) return true;
            }
            catch { }
            return false;
        }

        private static List<DashboardFilterBinding> ReadDefinitions(Workbook workbook)
        {
            var result = new List<DashboardFilterBinding>();
            for (var index = 1; index <= workbook.Names.Count; index++)
            {
                dynamic name = workbook.Names.Item(index);
                var fullName = Convert.ToString(name.Name);
                var separator = fullName.LastIndexOf('!');
                var shortName = separator >= 0 ? fullName.Substring(separator + 1) : fullName;
                if (!shortName.StartsWith(DefinitionPrefix, StringComparison.OrdinalIgnoreCase)) continue;
                Range metadata;
                try { metadata = name.RefersToRange; }
                catch { continue; }
                if (!string.Equals(Convert.ToString(((Range)metadata.Cells[1, 1]).Value2), DefinitionMarker,
                    StringComparison.Ordinal)) continue;

                for (var row = 2; row <= metadata.Rows.Count; row++)
                {
                    var dashboardSheet = Convert.ToString(((Range)metadata.Cells[row, 1]).Value2);
                    var supportSheet = Convert.ToString(((Range)metadata.Cells[row, 2]).Value2);
                    var controlAddress = Convert.ToString(((Range)metadata.Cells[row, 3]).Value2);
                    var fieldName = Convert.ToString(((Range)metadata.Cells[row, 4]).Value2);
                    if (string.IsNullOrWhiteSpace(dashboardSheet) || string.IsNullOrWhiteSpace(controlAddress) ||
                        string.IsNullOrWhiteSpace(fieldName)) continue;
                    result.Add(new DashboardFilterBinding
                    {
                        DashboardSheet = dashboardSheet,
                        SupportSheet = supportSheet,
                        ControlAddress = controlAddress,
                        FieldName = fieldName,
                        ListName = Convert.ToString(((Range)metadata.Cells[row, 5]).Value2)
                    });
                }
            }
            return result;
        }

        private static void Try(System.Action action)
        {
            try { action(); } catch { }
        }
    }
}
