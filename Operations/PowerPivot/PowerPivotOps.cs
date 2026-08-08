using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using System.Threading;
using Microsoft.Office.Interop.Excel;

namespace AgentForExcel.Operations.PowerPivot
{
    internal sealed class ModelFieldReference
    {
        public string Table { get; set; }
        public string Field { get; set; }
    }

    internal static class PowerPivotSupport
    {
        internal const string ListPrefix = "__AGENT_MODEL_LIST__";

        internal static Workbook GetWorkbook(AppContext context)
        {
            var workbook = context?.Excel?.ActiveWorkbook;
            if (workbook == null) throw new InvalidOperationException("当前没有打开的工作簿。");
            return workbook;
        }

        internal static dynamic GetModel(Workbook workbook)
        {
            try
            {
                dynamic workbookDynamic = workbook;
                dynamic model = workbookDynamic.Model;
                model.Initialize();
                return model;
            }
            catch (Exception ex)
            {
                throw new NotSupportedException(
                    "当前 Excel 版本或文件格式不支持 Power Pivot 数据模型。请使用 Microsoft 365/Excel 2016 及以上桌面版，并保存为 .xlsx 或 .xlsm。", ex);
            }
        }

        internal static dynamic FindModelTable(Workbook workbook, dynamic model, string tableName)
        {
            dynamic tables = model.ModelTables;
            for (var index = 1; index <= Convert.ToInt32(tables.Count); index++)
            {
                dynamic table = tables.Item(index);
                if (string.Equals(Convert.ToString(table.Name), tableName, StringComparison.OrdinalIgnoreCase))
                    return table;
            }
            var mappedName = ReadModelAlias(workbook, tableName);
            if (!string.IsNullOrWhiteSpace(mappedName))
            {
                for (var index = 1; index <= Convert.ToInt32(tables.Count); index++)
                {
                    dynamic table = tables.Item(index);
                    if (string.Equals(Convert.ToString(table.Name), mappedName, StringComparison.OrdinalIgnoreCase))
                        return table;
                }
            }
            var expectedConnection = "Agent Model - " + tableName;
            for (var index = 1; index <= Convert.ToInt32(tables.Count); index++)
            {
                dynamic table = tables.Item(index);
                try
                {
                    if ((Convert.ToString(table.SourceWorkbookConnection.Name) ?? string.Empty).StartsWith(
                        expectedConnection, StringComparison.OrdinalIgnoreCase)) return table;
                }
                catch { }
            }
            return null;
        }

        private static string AliasStorageName(string alias)
        {
            unchecked
            {
                uint hash = 2166136261;
                foreach (var character in alias ?? string.Empty)
                {
                    hash ^= character;
                    hash *= 16777619;
                }
                return "__AgentModelAlias_" + hash.ToString("X8");
            }
        }

        private static void SaveModelAlias(Workbook workbook, string alias, string actualTableName)
        {
            var storageName = AliasStorageName(alias);
            try { workbook.Names.Item(storageName).Delete(); } catch { }
            var escaped = (actualTableName ?? string.Empty).Replace("\"", "\"\"");
            workbook.Names.Add(Name: storageName, RefersTo: "=\"" + escaped + "\"", Visible: false);
        }

        private static string ReadModelAlias(Workbook workbook, string alias)
        {
            try
            {
                var refersTo = Convert.ToString(workbook.Names.Item(AliasStorageName(alias)).RefersTo) ?? string.Empty;
                if (refersTo.StartsWith("=\"", StringComparison.Ordinal) && refersTo.EndsWith("\"", StringComparison.Ordinal))
                    return refersTo.Substring(2, refersTo.Length - 3).Replace("\"\"", "\"");
            }
            catch { }
            return null;
        }

        internal static dynamic FindMeasure(dynamic model, string measureName)
        {
            dynamic measures = model.ModelMeasures;
            for (var index = 1; index <= Convert.ToInt32(measures.Count); index++)
            {
                dynamic measure = measures.Item(index);
                if (string.Equals(Convert.ToString(measure.Name), measureName, StringComparison.OrdinalIgnoreCase))
                    return measure;
            }
            return null;
        }

        internal static dynamic AddQueryToModel(Workbook workbook, string queryName)
        {
            dynamic workbookDynamic = workbook;
            dynamic queries = workbookDynamic.Queries;
            dynamic query = null;
            for (var index = 1; index <= Convert.ToInt32(queries.Count); index++)
            {
                dynamic candidate = queries.Item(index);
                if (string.Equals(Convert.ToString(candidate.Name), queryName, StringComparison.OrdinalIgnoreCase))
                {
                    query = candidate;
                    break;
                }
            }
            if (query == null) throw new ArgumentException("找不到 Power Query 查询「" + queryName + "」。");

            dynamic model = workbookDynamic.Model;
            model.Initialize();

            var connectionName = "Agent Model - " + queryName;
            dynamic connections = workbookDynamic.Connections;
            for (var index = 1; index <= Convert.ToInt32(connections.Count); index++)
            {
                dynamic connection = connections.Item(index);
                if (string.Equals(Convert.ToString(connection.Name), connectionName, StringComparison.OrdinalIgnoreCase))
                {
                    try { connection.Delete(); } catch { }
                    break;
                }
            }
            var modelQueryName = PrepareModelQuery(workbook, queryName);
            var connectionString = "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" +
                                   modelQueryName + ";Extended Properties=\"\"";
            dynamic modelConnection = CreateModelConnection(connections, connectionName, modelQueryName, connectionString);
            Exception refreshError = null;
            try { modelConnection.Refresh(); } catch (Exception ex) { refreshError = ex; }
            try { workbook.Application.CalculateUntilAsyncQueriesDone(); } catch { }
            model.Refresh();
            dynamic modelTable = WaitForModelTable(workbook, model, queryName);
            if (modelTable == null)
            {
                if (refreshError != null)
                    throw new InvalidOperationException("Power Pivot connection refresh failed: " + refreshError.Message, refreshError);
                var names = new List<string>();
                for (var connectionIndex = 1; connectionIndex <= Convert.ToInt32(connections.Count); connectionIndex++)
                {
                    dynamic connection = connections.Item(connectionIndex);
                    var details = Convert.ToString(connection.Name);
                    try { details += "{model=" + Convert.ToString(connection.ModelConnection) + "}"; } catch { }
                    try { details += "{command=" + Convert.ToString(connection.OLEDBConnection.CommandText) + "}"; } catch { }
                    names.Add("connection=" + details);
                }
                dynamic tables = model.ModelTables;
                for (var index = 1; index <= Convert.ToInt32(tables.Count); index++)
                {
                    dynamic item = tables.Item(index);
                    var columns = new List<string>();
                    try
                    {
                        dynamic itemColumns = item.ModelTableColumns;
                        for (var columnIndex = 1; columnIndex <= Convert.ToInt32(itemColumns.Count); columnIndex++)
                            columns.Add(Convert.ToString(itemColumns.Item(columnIndex).Name));
                    }
                    catch { }
                    var source = string.Empty;
                    try { source = Convert.ToString(item.SourceWorkbookConnection.Name); } catch { }
                    var rows = string.Empty;
                    try { rows = Convert.ToString(item.RecordCount); } catch { }
                    names.Add(Convert.ToString(item.Name) + "[" + string.Join(",", columns) + "; rows=" + rows + "; source=" + source + "]");
                }
                throw new InvalidOperationException("查询连接已创建，但 Excel 没有把「" + queryName +
                    "」载入数据模型。当前模型表：" + (names.Count == 0 ? "(空)" : string.Join("、", names)) + "。");
            }
            SaveModelAlias(workbook, queryName, Convert.ToString(modelTable.Name));
            return modelTable;
        }

        private static dynamic WaitForModelTable(Workbook workbook, dynamic model, string queryName)
        {
            for (var attempt = 0; attempt < 8; attempt++)
            {
                try { workbook.Application.CalculateUntilAsyncQueriesDone(); } catch { }
                dynamic table = FindModelTable(workbook, model, queryName);
                if (table != null) return table;
                table = FindModelTableByMaterializedSchema(workbook, model, queryName);
                if (table != null) return table;
                Thread.Sleep(250);
            }
            return null;
        }

        private static dynamic FindModelTableByMaterializedSchema(Workbook workbook, dynamic model, string queryName)
        {
            var expectedColumns = new List<string>();
            var expectedRows = -1;
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                dynamic listObjects = sheet.ListObjects;
                for (var index = 1; index <= Convert.ToInt32(listObjects.Count); index++)
                {
                    dynamic listObject = listObjects.Item(index);
                    if (!PowerQuery.PowerQuerySupport.IsFallbackTableForQuery(listObject, queryName)) continue;
                    dynamic listColumns = listObject.ListColumns;
                    for (var columnIndex = 1; columnIndex <= Convert.ToInt32(listColumns.Count); columnIndex++)
                        expectedColumns.Add(Convert.ToString(listColumns.Item(columnIndex).Name));
                    try { expectedRows = Convert.ToInt32(listObject.DataBodyRange.Rows.Count); } catch { expectedRows = 0; }
                    break;
                }
                if (expectedColumns.Count > 0) break;
            }
            if (expectedColumns.Count == 0) return null;

            var matches = new List<dynamic>();
            dynamic tables = model.ModelTables;
            for (var tableIndex = 1; tableIndex <= Convert.ToInt32(tables.Count); tableIndex++)
            {
                dynamic candidate = tables.Item(tableIndex);
                dynamic columns = candidate.ModelTableColumns;
                if (Convert.ToInt32(columns.Count) != expectedColumns.Count) continue;
                var same = true;
                for (var columnIndex = 1; columnIndex <= expectedColumns.Count; columnIndex++)
                {
                    if (!string.Equals(Convert.ToString(columns.Item(columnIndex).Name), expectedColumns[columnIndex - 1],
                        StringComparison.OrdinalIgnoreCase))
                    {
                        same = false;
                        break;
                    }
                }
                if (!same) continue;
                try
                {
                    var recordCount = Convert.ToInt32(candidate.RecordCount);
                    if (expectedRows >= 0 && recordCount != expectedRows) continue;
                }
                catch { }
                matches.Add(candidate);
            }
            return matches.Count == 1 ? matches[0] : null;
        }

        private static dynamic CreateModelConnection(dynamic connections, string connectionName, string queryName, string connectionString)
        {
            // Excel 为 Mashup/Power Query 生成的数据模型连接，会把查询标识作为
            // 带双引号的 xlCmdTableCollection 命令文本传入。测试宿主必须与
            // Excel 位宽一致；跨位宽调用会被 Mashup 解析成 "default" 假表。
            var commandText = "\"" + queryName.Replace("\"", "\"\"") + "\"";
            return connections.Add2(
                connectionName,
                "Agent for Excel 数据模型连接 - " + queryName,
                connectionString,
                commandText,
                XlCmdType.xlCmdTableCollection,
                true,
                false);
        }

        private static string PrepareModelQuery(Workbook workbook, string queryName)
        {
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                dynamic listObjects = sheet.ListObjects;
                for (var index = 1; index <= Convert.ToInt32(listObjects.Count); index++)
                {
                    dynamic listObject = listObjects.Item(index);
                    if (!PowerQuery.PowerQuerySupport.IsFallbackTableForQuery(listObject, queryName)) continue;
                    var tableName = Convert.ToString(listObject.Name).Replace("\"", "\"\"");
                    var formula = "let\n    Source = Excel.CurrentWorkbook(){[Name=\"" + tableName + "\"]}[Content]\nin\n    Source";
                    var bridgeName = "AgentModel_" + AliasStorageName(queryName).Substring("__AgentModelAlias_".Length);
                    dynamic queries = ((dynamic)workbook).Queries;
                    dynamic bridge = null;
                    for (var queryIndex = 1; queryIndex <= Convert.ToInt32(queries.Count); queryIndex++)
                    {
                        dynamic candidate = queries.Item(queryIndex);
                        if (string.Equals(Convert.ToString(candidate.Name), bridgeName, StringComparison.OrdinalIgnoreCase))
                        {
                            bridge = candidate;
                            break;
                        }
                    }
                    if (bridge == null) queries.Add(bridgeName, formula, "Agent for Excel model bridge for " + queryName);
                    else bridge.Formula = formula;
                    return bridgeName;
                }
            }
            return queryName;
        }

        internal static dynamic FindColumn(dynamic modelTable, string columnName)
        {
            dynamic columns = modelTable.ModelTableColumns;
            for (var index = 1; index <= Convert.ToInt32(columns.Count); index++)
            {
                dynamic column = columns.Item(index);
                if (string.Equals(Convert.ToString(column.Name), columnName, StringComparison.OrdinalIgnoreCase))
                    return column;
            }
            throw new ArgumentException("数据模型表「" + Convert.ToString(modelTable.Name) + "」中不存在字段「" + columnName + "」。");
        }

        internal static dynamic GetFormat(dynamic model, string format, int decimalPlaces, string currencySymbol)
        {
            switch ((format ?? "general").Trim().ToLowerInvariant())
            {
                case "currency":
                    dynamic currency = model.ModelFormatCurrency;
                    currency.DecimalPlaces = decimalPlaces;
                    if (!string.IsNullOrWhiteSpace(currencySymbol)) currency.Symbol = currencySymbol;
                    return currency;
                case "whole_number": return model.ModelFormatWholeNumber;
                case "decimal":
                    dynamic number = model.ModelFormatDecimalNumber;
                    number.DecimalPlaces = decimalPlaces;
                    return number;
                case "percentage":
                    dynamic percentage = model.ModelFormatPercentageNumber;
                    percentage.DecimalPlaces = decimalPlaces;
                    return percentage;
                case "date": return model.ModelFormatDate;
                case "scientific": return model.ModelFormatScientificNumber;
                case "boolean": return model.ModelFormatBoolean;
                case "general": return model.ModelFormatGeneral;
                default: throw new ArgumentException("不支持的度量值格式：" + format);
            }
        }

        internal static void ValidateDax(string formula)
        {
            if (string.IsNullOrWhiteSpace(formula)) throw new ArgumentException("DAX 公式不能为空。");
            if (formula.Length > 8000) throw new ArgumentException("DAX 公式过长，最多允许 8000 个字符。");
            var upper = formula.ToUpperInvariant();
            var forbidden = new[] { "EVALUATE ", "DEFINE ", "CREATE ", "ALTER ", "DROP " };
            if (forbidden.Any(token => upper.Contains(token)))
                throw new ArgumentException("这里只允许度量值表达式，不允许 DAX 查询或模型修改命令。");
        }

        internal static string RewriteTableReference(string formula, string requestedTable, string actualTable)
        {
            if (string.Equals(requestedTable, actualTable, StringComparison.OrdinalIgnoreCase)) return formula;
            var replacement = "'" + actualTable.Replace("'", "''") + "'[";
            return formula
                .Replace("'" + requestedTable.Replace("'", "''") + "'[", replacement)
                .Replace(requestedTable + "[", replacement);
        }

        internal static string CubeFieldName(string table, string field) =>
            "[" + table.Replace("]", "]]" ) + "].[" + field.Replace("]", "]]" ) + "]";
    }

    public sealed class ListModelOp : IOperation
    {
        public string ToolName => "pp_list_model";
        public bool IsWriteOperation => false;
        public string Describe() => "读取当前工作簿的 Power Pivot 数据模型";

        public string Execute(AppContext context)
        {
            var workbook = PowerPivotSupport.GetWorkbook(context);
            dynamic model = PowerPivotSupport.GetModel(workbook);
            var tables = new List<object>();
            dynamic modelTables = model.ModelTables;
            for (var index = 1; index <= Convert.ToInt32(modelTables.Count); index++)
            {
                dynamic table = modelTables.Item(index);
                var columns = new List<string>();
                dynamic modelColumns = table.ModelTableColumns;
                for (var columnIndex = 1; columnIndex <= Convert.ToInt32(modelColumns.Count); columnIndex++)
                    columns.Add(Convert.ToString(modelColumns.Item(columnIndex).Name));
                tables.Add(new { name = Convert.ToString(table.Name), rows = Convert.ToInt32(table.RecordCount), columns });
            }

            var relationships = new List<object>();
            dynamic modelRelationships = model.ModelRelationships;
            for (var index = 1; index <= Convert.ToInt32(modelRelationships.Count); index++)
            {
                dynamic relationship = modelRelationships.Item(index);
                relationships.Add(new
                {
                    from_table = Convert.ToString(relationship.ForeignKeyTable.Name),
                    from_column = Convert.ToString(relationship.ForeignKeyColumn.Name),
                    to_table = Convert.ToString(relationship.PrimaryKeyTable.Name),
                    to_column = Convert.ToString(relationship.PrimaryKeyColumn.Name)
                });
            }

            var measures = new List<object>();
            dynamic modelMeasures = model.ModelMeasures;
            for (var index = 1; index <= Convert.ToInt32(modelMeasures.Count); index++)
            {
                dynamic measure = modelMeasures.Item(index);
                measures.Add(new
                {
                    name = Convert.ToString(measure.Name),
                    table = Convert.ToString(measure.AssociatedTable.Name),
                    formula = Convert.ToString(measure.Formula),
                    description = Convert.ToString(measure.Description)
                });
            }
            return PowerPivotSupport.ListPrefix + JsonSerializer.Serialize(new { tables, relationships, measures });
        }

        public sealed class Factory : IOperationFactory
        {
            public string ToolName => "pp_list_model";
            public IOperation Parse(string argumentsJson) => new ListModelOp();
        }
    }

    public sealed class AddQueryToModelOp : IOperation
    {
        private readonly string _queryName;
        private AddQueryToModelOp(string queryName) { _queryName = queryName; }
        public string ToolName => "pp_add_query_to_model";
        public bool IsWriteOperation => true;
        public string Describe() => "把 Power Query 查询「" + _queryName + "」加入 Power Pivot 数据模型";
        public string Execute(AppContext context)
        {
            var workbook = PowerPivotSupport.GetWorkbook(context);
            dynamic table = PowerPivotSupport.AddQueryToModel(workbook, _queryName);
            return "已将查询「" + _queryName + "」加入数据模型，模型表包含 " +
                   Convert.ToInt32(table.RecordCount) + " 行、" + Convert.ToInt32(table.ModelTableColumns.Count) + " 列。";
        }
        public sealed class Factory : IOperationFactory
        {
            public string ToolName => "pp_add_query_to_model";
            public IOperation Parse(string argumentsJson) => new AddQueryToModelOp(ReadRequired(argumentsJson, "query_name"));
        }
        internal static string ReadRequired(string json, string property)
        {
            using (var document = JsonDocument.Parse(json))
            {
                if (!document.RootElement.TryGetProperty(property, out var value) || value.ValueKind != JsonValueKind.String ||
                    string.IsNullOrWhiteSpace(value.GetString())) throw new ArgumentException(property + " 不能为空。");
                return value.GetString().Trim();
            }
        }
    }

    public sealed class RefreshModelOp : IOperation
    {
        public string ToolName => "pp_refresh_model";
        public bool IsWriteOperation => true;
        public string Describe() => "刷新 Power Pivot 数据模型及其模型透视表";

        public string Execute(AppContext context)
        {
            var workbook = PowerPivotSupport.GetWorkbook(context);
            dynamic model = PowerPivotSupport.GetModel(workbook);
            try { workbook.RefreshAll(); } catch { }
            try { workbook.Application.CalculateUntilAsyncQueriesDone(); } catch { }
            model.Refresh();
            try { workbook.Application.CalculateUntilAsyncQueriesDone(); } catch { }

            var pivotCount = 0;
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                dynamic pivots = sheet.PivotTables();
                for (var index = 1; index <= Convert.ToInt32(pivots.Count); index++)
                {
                    dynamic pivot = pivots.Item(index);
                    try
                    {
                        if (Convert.ToBoolean(pivot.PivotCache().IsConnected))
                        {
                            pivot.RefreshTable();
                            pivotCount++;
                        }
                    }
                    catch { }
                }
            }
            return "已刷新 Power Pivot 数据模型（" + Convert.ToInt32(model.ModelTables.Count) +
                   " 张模型表）及 " + pivotCount + " 张模型透视表。";
        }

        public sealed class Factory : IOperationFactory
        {
            public string ToolName => "pp_refresh_model";
            public IOperation Parse(string argumentsJson) => new RefreshModelOp();
        }
    }

    public sealed class AddRelationshipOp : IOperation
    {
        private readonly string _fromTable, _fromColumn, _toTable, _toColumn;
        private AddRelationshipOp(string fromTable, string fromColumn, string toTable, string toColumn)
        { _fromTable = fromTable; _fromColumn = fromColumn; _toTable = toTable; _toColumn = toColumn; }
        public string ToolName => "pp_add_relationship";
        public bool IsWriteOperation => true;
        public string Describe() => "建立模型关系 " + _fromTable + "[" + _fromColumn + "] → " + _toTable + "[" + _toColumn + "]";
        public string Execute(AppContext context)
        {
            var workbook = PowerPivotSupport.GetWorkbook(context);
            dynamic model = PowerPivotSupport.GetModel(workbook);
            dynamic fromTable = PowerPivotSupport.FindModelTable(workbook, model, _fromTable);
            dynamic toTable = PowerPivotSupport.FindModelTable(workbook, model, _toTable);
            if (fromTable == null || toTable == null) throw new ArgumentException("关系两端的模型表不存在，请先加入数据模型。");
            dynamic fromColumn = PowerPivotSupport.FindColumn(fromTable, _fromColumn);
            dynamic toColumn = PowerPivotSupport.FindColumn(toTable, _toColumn);
            model.ModelRelationships.Add(fromColumn, toColumn);
            return "已建立一对多关系：" + _fromTable + "[" + _fromColumn + "] → " + _toTable + "[" + _toColumn + "]。";
        }
        public sealed class Factory : IOperationFactory
        {
            public string ToolName => "pp_add_relationship";
            public IOperation Parse(string argumentsJson)
            {
                using (var d = JsonDocument.Parse(argumentsJson))
                {
                    return new AddRelationshipOp(Read(d.RootElement, "from_table"), Read(d.RootElement, "from_column"),
                        Read(d.RootElement, "to_table"), Read(d.RootElement, "to_column"));
                }
            }
            private static string Read(JsonElement root, string name)
            {
                if (!root.TryGetProperty(name, out var value) || value.ValueKind != JsonValueKind.String || string.IsNullOrWhiteSpace(value.GetString()))
                    throw new ArgumentException(name + " 不能为空。");
                return value.GetString().Trim();
            }
        }
    }

    public sealed class AddMeasureOp : IOperation
    {
        private readonly string _table, _name, _formula, _format, _symbol, _description;
        private readonly int _decimalPlaces;
        private readonly bool _replace;
        private AddMeasureOp(string table, string name, string formula, string format, int decimals,
            string symbol, string description, bool replace)
        { _table = table; _name = name; _formula = formula; _format = format; _decimalPlaces = decimals; _symbol = symbol; _description = description; _replace = replace; }
        public string ToolName => "pp_add_measure";
        public bool IsWriteOperation => true;
        public string Describe() => "在模型表「" + _table + "」创建 DAX 度量值「" + _name + "」";
        public string Execute(AppContext context)
        {
            PowerPivotSupport.ValidateDax(_formula);
            var workbook = PowerPivotSupport.GetWorkbook(context);
            dynamic model = PowerPivotSupport.GetModel(workbook);
            dynamic table = PowerPivotSupport.FindModelTable(workbook, model, _table);
            if (table == null) throw new ArgumentException("数据模型中不存在表「" + _table + "」。");
            dynamic existing = PowerPivotSupport.FindMeasure(model, _name);
            if (existing != null && !_replace) throw new InvalidOperationException("度量值「" + _name + "」已经存在。");
            if (existing != null) existing.Delete();
            var dax = _formula.TrimStart().StartsWith("=", StringComparison.Ordinal) ? _formula.TrimStart().Substring(1) : _formula;
            dax = PowerPivotSupport.RewriteTableReference(dax, _table, Convert.ToString(table.Name));
            dynamic format = PowerPivotSupport.GetFormat(model, _format, _decimalPlaces, _symbol);
            dynamic measure = model.ModelMeasures.Add(_name, table, dax, format, _description ?? string.Empty);
            return "已创建 DAX 度量值「" + Convert.ToString(measure.Name) + "」：" + dax + "。";
        }
        public sealed class Factory : IOperationFactory
        {
            public string ToolName => "pp_add_measure";
            public IOperation Parse(string argumentsJson)
            {
                using (var d = JsonDocument.Parse(argumentsJson))
                {
                    var r = d.RootElement;
                    return new AddMeasureOp(Read(r, "table"), Read(r, "measure_name"), Read(r, "formula"),
                        Optional(r, "format") ?? "general", Int(r, "decimal_places", 2), Optional(r, "currency_symbol"),
                        Optional(r, "description"), Bool(r, "replace_existing", false));
                }
            }
            private static string Read(JsonElement r, string n) => Optional(r, n) ?? throw new ArgumentException(n + " 不能为空。");
            private static string Optional(JsonElement r, string n) => r.TryGetProperty(n, out var v) && v.ValueKind == JsonValueKind.String && !string.IsNullOrWhiteSpace(v.GetString()) ? v.GetString().Trim() : null;
            private static int Int(JsonElement r, string n, int f) => r.TryGetProperty(n, out var v) && v.ValueKind == JsonValueKind.Number ? v.GetInt32() : f;
            private static bool Bool(JsonElement r, string n, bool f) => r.TryGetProperty(n, out var v) && (v.ValueKind == JsonValueKind.True || v.ValueKind == JsonValueKind.False) ? v.GetBoolean() : f;
        }
    }

    public sealed class CreateModelPivotOp : IOperation
    {
        private readonly string _sheetName, _address, _name;
        private readonly ModelFieldReference[] _rows, _columns, _filters;
        private readonly string[] _measures;
        private CreateModelPivotOp(string sheetName, string address, string name, ModelFieldReference[] rows,
            ModelFieldReference[] columns, ModelFieldReference[] filters, string[] measures)
        { _sheetName = sheetName; _address = address; _name = name; _rows = rows; _columns = columns; _filters = filters; _measures = measures; }
        public string ToolName => "pp_create_model_pivot";
        public bool IsWriteOperation => true;
        public string Describe() => "基于 Power Pivot 模型创建透视表「" + _name + "」";
        public string Execute(AppContext context)
        {
            var workbook = PowerPivotSupport.GetWorkbook(context);
            dynamic model = PowerPivotSupport.GetModel(workbook);
            Worksheet sheet = null;
            try
            {
                sheet = Analysis.AnalysisSheetSupport.CreateUniqueWorksheet(context, _sheetName);
                dynamic workbookDynamic = workbook;
                dynamic cache = workbookDynamic.PivotCaches().Create(
                    XlPivotTableSourceType.xlExternal, model.DataModelConnection, XlPivotTableVersionList.xlPivotTableVersion15);
                dynamic pivot = cache.CreatePivotTable(sheet.Range[_address], _name);
                ApplyFields(workbook, model, pivot, _rows, XlPivotFieldOrientation.xlRowField);
                ApplyFields(workbook, model, pivot, _columns, XlPivotFieldOrientation.xlColumnField);
                ApplyFields(workbook, model, pivot, _filters, XlPivotFieldOrientation.xlPageField);
                foreach (var measure in _measures)
                {
                    dynamic field = pivot.CubeFields.Item("[Measures].[" + measure.Replace("]", "]]" ) + "]");
                    field.Orientation = XlPivotFieldOrientation.xlDataField;
                }
                pivot.TableStyle2 = "PivotStyleMedium9";
                pivot.RowAxisLayout(XlLayoutRowType.xlTabularRow);
                sheet.Columns.AutoFit();
                sheet.Activate();
                return "已创建模型透视表「" + _name + "」，位置为 " + sheet.Name + "!" + _address +
                       "，包含 " + _measures.Length + " 个 DAX 度量值。";
            }
            catch
            {
                if (sheet != null) Analysis.AnalysisSheetSupport.DeleteWorksheetSilently(context, sheet);
                throw;
            }
        }
        private static void ApplyFields(Workbook workbook, dynamic model, dynamic pivot, IEnumerable<ModelFieldReference> fields, XlPivotFieldOrientation orientation)
        {
            var position = 1;
            foreach (var reference in fields)
            {
                dynamic modelTable = PowerPivotSupport.FindModelTable(workbook, model, reference.Table);
                if (modelTable == null) throw new ArgumentException("数据模型中不存在表「" + reference.Table + "」。");
                dynamic field = pivot.CubeFields.Item(PowerPivotSupport.CubeFieldName(
                    Convert.ToString(modelTable.Name), reference.Field));
                field.Orientation = orientation;
                field.Position = position++;
            }
        }
        public sealed class Factory : IOperationFactory
        {
            public string ToolName => "pp_create_model_pivot";
            public IOperation Parse(string argumentsJson)
            {
                using (var d = JsonDocument.Parse(argumentsJson))
                {
                    var r = d.RootElement;
                    var measures = Strings(r, "measures");
                    if (measures.Length == 0) throw new ArgumentException("measures 至少需要一个度量值。");
                    return new CreateModelPivotOp(Optional(r, "destination_sheet") ?? "Agent模型透视",
                        Optional(r, "destination_address") ?? "A1", Optional(r, "name") ?? "AgentModelPivot",
                        Fields(r, "rows"), Fields(r, "columns"), Fields(r, "filters"), measures);
                }
            }
            private static string Optional(JsonElement r, string n) => r.TryGetProperty(n, out var v) && v.ValueKind == JsonValueKind.String && !string.IsNullOrWhiteSpace(v.GetString()) ? v.GetString().Trim() : null;
            private static string[] Strings(JsonElement r, string n) => !r.TryGetProperty(n, out var v) ? new string[0] : v.EnumerateArray().Select(x => x.GetString()?.Trim()).Where(x => !string.IsNullOrWhiteSpace(x)).ToArray();
            private static ModelFieldReference[] Fields(JsonElement r, string n)
            {
                if (!r.TryGetProperty(n, out var v)) return new ModelFieldReference[0];
                return v.EnumerateArray().Select(x => new ModelFieldReference
                {
                    Table = Optional(x, "table") ?? throw new ArgumentException(n + ".table 不能为空。"),
                    Field = Optional(x, "field") ?? throw new ArgumentException(n + ".field 不能为空。")
                }).ToArray();
            }
        }
    }
}
