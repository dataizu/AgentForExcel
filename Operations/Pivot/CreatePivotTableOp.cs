using System;
using System.Collections.Generic;
using System.Text.Json;
using Microsoft.Office.Interop.Excel;

namespace AgentForExcel.Operations.Pivot
{
    /// <summary>根据标准二维表创建普通（非 OLAP）数据透视表。</summary>
    public sealed class CreatePivotTableOp : IOperation
    {
        public string ToolName => "pivot_create";
        public bool IsWriteOperation => true;

        private readonly string _sourceSheet;
        private readonly string _sourceAddress;
        private readonly string _destinationSheet;
        private readonly string _destinationAddress;
        private readonly string _name;
        private readonly string[] _rowFields;
        private readonly string[] _columnFields;
        private readonly string[] _filterFields;
        private readonly ValueSpec[] _values;

        private CreatePivotTableOp(
            string sourceSheet, string sourceAddress, string destinationSheet, string destinationAddress,
            string name, string[] rowFields, string[] columnFields, string[] filterFields, ValueSpec[] values)
        {
            _sourceSheet = sourceSheet;
            _sourceAddress = sourceAddress;
            _destinationSheet = destinationSheet;
            _destinationAddress = destinationAddress;
            _name = name;
            _rowFields = rowFields;
            _columnFields = columnFields;
            _filterFields = filterFields;
            _values = values;
        }

        public string Describe()
        {
            var source = (string.IsNullOrWhiteSpace(_sourceSheet) ? "活动工作表" : "工作表「" + _sourceSheet + "」") + "!" + _sourceAddress;
            var destination = string.IsNullOrWhiteSpace(_destinationSheet) ? "新工作表「透视分析」" : "工作表「" + _destinationSheet + "」";
            return $"根据 {source} 创建数据透视表「{_name}」，输出到 {destination}!{_destinationAddress}";
        }

        public string Execute(AppContext context)
        {
            var workbook = context.Excel.ActiveWorkbook;
            if (workbook == null) throw new InvalidOperationException("当前没有打开的工作簿。");

            var sourceSheet = Cell.CellOperationSupport.GetWorksheet(context, _sourceSheet);
            var sourceRange = Cell.CellOperationSupport.GetRange(sourceSheet, _sourceAddress);
            if (sourceRange.Rows.Count < 2 || sourceRange.Columns.Count < 2)
                throw new ArgumentException("透视表源区域至少需要 2 行 × 2 列，并包含标题行。");

            ValidateFields(sourceRange);
            var destinationSheet = GetOrCreateDestinationSheet(workbook, _destinationSheet);
            var destination = Cell.CellOperationSupport.GetRange(destinationSheet, _destinationAddress);
            if (destination.Value2 != null)
                throw new InvalidOperationException($"目标位置 {destinationSheet.Name}!{_destinationAddress} 不是空白单元格。");

            var pivotName = MakeUniquePivotName(workbook, _name);
            var caches = workbook.PivotCaches();
            var cache = caches.Create(XlPivotTableSourceType.xlDatabase, sourceRange, Type.Missing);
            var pivot = cache.CreatePivotTable(destination, pivotName, Type.Missing, Type.Missing);

            ApplyFields(pivot, _rowFields, XlPivotFieldOrientation.xlRowField);
            ApplyFields(pivot, _columnFields, XlPivotFieldOrientation.xlColumnField);
            ApplyFields(pivot, _filterFields, XlPivotFieldOrientation.xlPageField);

            foreach (var value in _values)
            {
                var field = (PivotField)pivot.PivotFields(value.Field);
                var caption = string.IsNullOrWhiteSpace(value.Caption)
                    ? FunctionCaption(value.Function) + value.Field
                    : value.Caption;
                pivot.AddDataField(field, caption, ParseFunction(value.Function));
            }

            try { pivot.RowAxisLayout(XlLayoutRowType.xlTabularRow); } catch { }
            pivot.TableStyle2 = "PivotStyleMedium2";
            pivot.RefreshTable();

            return $"已创建数据透视表「{pivot.Name}」，位置为 {destinationSheet.Name}!{destination.Address}，包含 {_values.Length} 个值字段。";
        }

        private void ValidateFields(Range sourceRange)
        {
            var available = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (var c = 1; c <= sourceRange.Columns.Count; c++)
            {
                var cell = (Range)sourceRange.Cells[1, c];
                var header = Convert.ToString(cell.Value2)?.Trim();
                if (string.IsNullOrEmpty(header)) throw new ArgumentException("透视表源区域的标题行不能包含空白列名。");
                if (!available.Add(header)) throw new ArgumentException("透视表源区域包含重复列名：「" + header + "」。");
            }

            foreach (var field in EnumerateRequestedFields())
                if (!available.Contains(field))
                    throw new ArgumentException("源区域中找不到字段「" + field + "」。");
        }

        private IEnumerable<string> EnumerateRequestedFields()
        {
            foreach (var field in _rowFields) yield return field;
            foreach (var field in _columnFields) yield return field;
            foreach (var field in _filterFields) yield return field;
            foreach (var value in _values) yield return value.Field;
        }

        private static void ApplyFields(PivotTable pivot, string[] fields, XlPivotFieldOrientation orientation)
        {
            for (var i = 0; i < fields.Length; i++)
            {
                var field = (PivotField)pivot.PivotFields(fields[i]);
                field.Orientation = orientation;
                field.Position = i + 1;
            }
        }

        private static Worksheet GetOrCreateDestinationSheet(Workbook workbook, string requestedName)
        {
            var name = string.IsNullOrWhiteSpace(requestedName) ? "透视分析" : requestedName.Trim();
            if (name.Length > 31) throw new ArgumentException("destination_sheet 不能超过 31 个字符。");

            try { return (Worksheet)workbook.Worksheets[name]; }
            catch
            {
                var last = workbook.Worksheets[workbook.Worksheets.Count];
                var sheet = (Worksheet)workbook.Worksheets.Add(Type.Missing, last, 1, XlSheetType.xlWorksheet);
                sheet.Name = name;
                return sheet;
            }
        }

        private static string MakeUniquePivotName(Workbook workbook, string preferredName)
        {
            var baseName = string.IsNullOrWhiteSpace(preferredName) ? "AgentPivot" : preferredName.Trim();
            var candidate = baseName;
            var suffix = 2;
            while (PivotNameExists(workbook, candidate)) candidate = baseName + suffix++;
            return candidate;
        }

        private static bool PivotNameExists(Workbook workbook, string name)
        {
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                var pivots = (PivotTables)sheet.PivotTables(Type.Missing);
                for (var i = 1; i <= pivots.Count; i++)
                    if (string.Equals(pivots.Item(i).Name, name, StringComparison.OrdinalIgnoreCase)) return true;
            }
            return false;
        }

        private static XlConsolidationFunction ParseFunction(string value)
        {
            switch ((value ?? "sum").Trim().ToLowerInvariant())
            {
                case "sum": return XlConsolidationFunction.xlSum;
                case "count": return XlConsolidationFunction.xlCount;
                case "average": return XlConsolidationFunction.xlAverage;
                case "max": return XlConsolidationFunction.xlMax;
                case "min": return XlConsolidationFunction.xlMin;
                default: throw new ArgumentException("values.function 仅支持 sum、count、average、max、min。");
            }
        }

        private static string FunctionCaption(string function)
        {
            switch ((function ?? "sum").Trim().ToLowerInvariant())
            {
                case "count": return "计数项:";
                case "average": return "平均值:";
                case "max": return "最大值:";
                case "min": return "最小值:";
                default: return "求和项:";
            }
        }

        private sealed class ValueSpec
        {
            public string Field { get; set; }
            public string Function { get; set; }
            public string Caption { get; set; }
        }

        public sealed class Factory : IOperationFactory
        {
            public string ToolName => "pivot_create";

            public IOperation Parse(string argumentsJson)
            {
                using (var doc = JsonDocument.Parse(argumentsJson))
                {
                    var root = doc.RootElement;
                    var sourceAddress = ReadString(root, "source_address");
                    if (string.IsNullOrWhiteSpace(sourceAddress)) throw new ArgumentException("source_address 不能为空。");

                    var values = ParseValues(root);
                    var rows = ReadStringArray(root, "rows");
                    var columns = ReadStringArray(root, "columns");
                    var filters = ReadStringArray(root, "filters");
                    if (values.Length == 0) throw new ArgumentException("values 至少需要一个值字段。");
                    if (rows.Length == 0 && columns.Length == 0)
                        throw new ArgumentException("rows 或 columns 至少需要一个分组字段。");

                    return new CreatePivotTableOp(
                        ReadString(root, "source_sheet"), sourceAddress,
                        ReadString(root, "destination_sheet"), ReadString(root, "destination_address") ?? "A1",
                        ReadString(root, "name") ?? "AgentPivot", rows, columns, filters, values);
                }
            }

            private static ValueSpec[] ParseValues(JsonElement root)
            {
                if (!root.TryGetProperty("values", out var valuesElement) || valuesElement.ValueKind != JsonValueKind.Array)
                    return new ValueSpec[0];

                var list = new List<ValueSpec>();
                foreach (var item in valuesElement.EnumerateArray())
                {
                    if (item.ValueKind != JsonValueKind.Object) throw new ArgumentException("values 每一项必须是对象。");
                    var field = ReadString(item, "field");
                    if (string.IsNullOrWhiteSpace(field)) throw new ArgumentException("values.field 不能为空。");
                    var function = ReadString(item, "function") ?? "sum";
                    ParseFunction(function);
                    list.Add(new ValueSpec { Field = field, Function = function, Caption = ReadString(item, "caption") });
                }
                return list.ToArray();
            }

            private static string[] ReadStringArray(JsonElement root, string name)
            {
                if (!root.TryGetProperty(name, out var element)) return new string[0];
                if (element.ValueKind != JsonValueKind.Array) throw new ArgumentException(name + " 必须是字符串数组。");
                var values = new List<string>();
                foreach (var item in element.EnumerateArray())
                {
                    if (item.ValueKind != JsonValueKind.String || string.IsNullOrWhiteSpace(item.GetString()))
                        throw new ArgumentException(name + " 不能包含空值。");
                    values.Add(item.GetString());
                }
                return values.ToArray();
            }

            private static string ReadString(JsonElement root, string name)
            {
                return root.TryGetProperty(name, out var value) && value.ValueKind == JsonValueKind.String
                    ? value.GetString()
                    : null;
            }
        }
    }
}
