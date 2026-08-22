using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.Json;
using Microsoft.Office.Interop.Excel;

namespace AgentForExcel.Operations.PowerQuery
{
    internal sealed class QueryColumnType
    {
        public string Field { get; set; }
        public string Type { get; set; }
    }

    internal sealed class QueryRename
    {
        public string From { get; set; }
        public string To { get; set; }
    }

    internal sealed class AgentPowerQueryMetadata
    {
        public string SourceSheet { get; set; }
        public string SourceAddress { get; set; }
        public string SourceName { get; set; }
        public bool RemoveBlankRows { get; set; }
        public bool TrimText { get; set; }
        public bool RemoveDuplicates { get; set; }
        public QueryRename[] Renames { get; set; }
        public QueryColumnType[] ColumnTypes { get; set; }
        public string[] SelectColumns { get; set; }
    }

    internal sealed class QueryLoadOutcome
    {
        public dynamic ListObject { get; set; }
        public bool UsedCompatibilityEngine { get; set; }
        /// <summary>兼容引擎按类型规则转换失败、已保留原值的单元格数。</summary>
        public int TypeCoercionFailures { get; set; }
    }

    internal static class PowerQuerySupport
    {
        internal const string MetadataPrefix = "__AGENT_PQ_META__";
        internal static Workbook GetWorkbook(AppContext context)
        {
            var workbook = context?.Excel?.ActiveWorkbook;
            if (workbook == null) throw new InvalidOperationException("当前没有打开的工作簿。");
            return workbook;
        }

        internal static dynamic FindQuery(Workbook workbook, string queryName)
        {
            dynamic workbookDynamic = workbook;
            dynamic queries = workbookDynamic.Queries;
            for (var index = 1; index <= Convert.ToInt32(queries.Count); index++)
            {
                dynamic query = queries.Item(index);
                if (string.Equals(Convert.ToString(query.Name), queryName, StringComparison.OrdinalIgnoreCase))
                    return query;
            }
            return null;
        }

        internal static string BuildRangeQueryFormula(
            string sourceName, bool removeBlankRows, bool trimText, bool removeDuplicates,
            IList<QueryRename> renames, IList<QueryColumnType> columnTypes, IList<string> selectColumns)
        {
            var lines = new List<string>();
            var previous = "Source";
            lines.Add("    Source = Excel.CurrentWorkbook(){[Name=\"" + EscapeM(sourceName) + "\"]}[Content]");
            AddStep(lines, ref previous, "Promoted Headers",
                "Table.PromoteHeaders(" + previous + ", [PromoteAllScalars=true])");
            if (removeBlankRows)
                AddStep(lines, ref previous, "Removed Blank Rows",
                    "Table.SelectRows(" + previous + ", each List.NonNullCount(Record.FieldValues(_)) > 0)");
            if (trimText)
                AddStep(lines, ref previous, "Trimmed Text",
                    "Table.TransformColumns(" + previous + ", List.Transform(Table.ColumnNames(" + previous + "), (columnName) => {columnName, each if _ is text then Text.Trim(_) else _, type any}))");
            if (renames != null && renames.Count > 0)
            {
                var pairs = string.Join(", ", renames.Select(rename =>
                    "{\"" + EscapeM(rename.From) + "\", \"" + EscapeM(rename.To) + "\"}"));
                AddStep(lines, ref previous, "Renamed Columns",
                    "Table.RenameColumns(" + previous + ", {" + pairs + "}, MissingField.Ignore)");
            }
            if (columnTypes != null && columnTypes.Count > 0)
            {
                var pairs = string.Join(", ", columnTypes.Select(item =>
                    "{\"" + EscapeM(item.Field) + "\", " + MapMType(item.Type) + "}"));
                AddStep(lines, ref previous, "Changed Type",
                    "Table.TransformColumnTypes(" + previous + ", {" + pairs + "}, \"zh-CN\")");
            }
            if (removeDuplicates)
                AddStep(lines, ref previous, "Removed Duplicates", "Table.Distinct(" + previous + ")");
            if (selectColumns != null && selectColumns.Count > 0)
            {
                var columns = string.Join(", ", selectColumns.Select(column => "\"" + EscapeM(column) + "\""));
                AddStep(lines, ref previous, "Selected Columns",
                    "Table.SelectColumns(" + previous + ", {" + columns + "}, MissingField.Ignore)");
            }
            return "let\n" + string.Join(",\n", lines) + "\nin\n    " + previous;
        }

        private static void AddStep(ICollection<string> lines, ref string previous, string stepName, string expression)
        {
            previous = "#\"" + stepName + "\"";
            lines.Add("    " + previous + " = " + expression);
        }

        internal static QueryLoadOutcome LoadQueryToSheet(
            Workbook workbook, Worksheet sheet, string queryName, Range destination, string tableName)
        {
            var connectionString = "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" +
                                   queryName + ";Extended Properties=\"\"";
            // Agent 自建查询包含可重复执行的结构化元数据。直接使用兼容物化引擎，
            // 避免部分 Excel 版本在原生 Mashup 工作表加载失败后污染整个会话的数据模型。
            dynamic agentQuery = FindQuery(workbook, queryName);
            AgentPowerQueryMetadata agentMetadata;
            if (TryReadMetadata(agentQuery, out agentMetadata))
            {
                int coercionFailures;
                dynamic fallbackTable = LoadFallbackTable(workbook, sheet, destination, queryName, agentMetadata, out coercionFailures);
                return new QueryLoadOutcome { ListObject = fallbackTable, UsedCompatibilityEngine = true, TypeCoercionFailures = coercionFailures };
            }
            try
            {
                dynamic listObjects = sheet.ListObjects;
                dynamic listObject = listObjects.Add(
                    XlListObjectSourceType.xlSrcExternal,
                    connectionString,
                    Type.Missing,
                    XlYesNoGuess.xlYes,
                    destination,
                    Type.Missing);
                dynamic queryTable = listObject.QueryTable;
                queryTable.CommandType = XlCmdType.xlCmdSql;
                queryTable.CommandText = new[] { "SELECT * FROM [" + queryName.Replace("]", "]]" ) + "]" };
                queryTable.RefreshStyle = XlCellInsertionMode.xlInsertDeleteCells;
                queryTable.BackgroundQuery = false;
                queryTable.Refresh(false);
                listObject.Name = MakeSafeTableName(tableName);
                return new QueryLoadOutcome { ListObject = listObject, UsedCompatibilityEngine = false };
            }
            catch (Exception nativeError)
            {
                DeleteAllListObjects(sheet);
                DeleteFailedQueryConnections(workbook, queryName);
                sheet.Cells.Clear();
                dynamic query = FindQuery(workbook, queryName);
                AgentPowerQueryMetadata metadata;
                if (!TryReadMetadata(query, out metadata))
                    throw new InvalidOperationException(
                        "Excel 原生 Power Query 加载失败，并且该查询不是 Agent 创建的兼容查询，无法自动回退。", nativeError);
                int coercionFailures;
                dynamic fallbackTable = LoadFallbackTable(workbook, sheet, destination, queryName, metadata, out coercionFailures);
                return new QueryLoadOutcome { ListObject = fallbackTable, UsedCompatibilityEngine = true, TypeCoercionFailures = coercionFailures };
            }
        }

        internal static string BuildDescription(string friendlyText, AgentPowerQueryMetadata metadata) =>
            friendlyText + "\n" + MetadataPrefix + JsonSerializer.Serialize(metadata);

        internal static string GetFriendlyDescription(string value)
        {
            var description = value ?? string.Empty;
            var index = description.IndexOf(MetadataPrefix, StringComparison.Ordinal);
            return index < 0 ? description : description.Substring(0, index).TrimEnd();
        }

        internal static bool TryReadMetadata(dynamic query, out AgentPowerQueryMetadata metadata)
        {
            metadata = null;
            if (query == null) return false;
            var description = Convert.ToString(query.Description) ?? string.Empty;
            var index = description.IndexOf(MetadataPrefix, StringComparison.Ordinal);
            if (index < 0) return false;
            try
            {
                metadata = JsonSerializer.Deserialize<AgentPowerQueryMetadata>(
                    description.Substring(index + MetadataPrefix.Length));
                return metadata != null && !string.IsNullOrWhiteSpace(metadata.SourceSheet) &&
                       !string.IsNullOrWhiteSpace(metadata.SourceAddress);
            }
            catch { return false; }
        }

        internal static dynamic RefreshFallbackTable(
            Workbook workbook, Worksheet sheet, dynamic oldTable, string queryName,
            out int typeCoercionFailures)
        {
            typeCoercionFailures = 0;
            dynamic query = FindQuery(workbook, queryName);
            AgentPowerQueryMetadata metadata;
            if (!TryReadMetadata(query, out metadata)) return null;
            dynamic oldRange = oldTable.Range;
            dynamic topLeft = oldRange.Cells.Item(1, 1);
            var destinationAddress = Convert.ToString(topLeft.Address);
            oldTable.Unlist();
            oldRange.Clear();
            return LoadFallbackTable(workbook, sheet, sheet.Range[destinationAddress], queryName, metadata,
                out typeCoercionFailures);
        }

        internal static bool IsFallbackTableForQuery(dynamic listObject, string queryName)
        {
            var prefix = MakeSafeTablePrefix(queryName);
            return Convert.ToString(listObject.Name).StartsWith(prefix, StringComparison.OrdinalIgnoreCase);
        }

        private static dynamic LoadFallbackTable(
            Workbook workbook, Worksheet destinationSheet, Range destination, string queryName,
            AgentPowerQueryMetadata metadata, out int typeCoercionFailures)
        {
            typeCoercionFailures = 0;
            dynamic workbookDynamic = workbook;
            var sourceSheet = (Worksheet)workbookDynamic.Worksheets.Item(metadata.SourceSheet);
            var sourceRange = sourceSheet.Range[metadata.SourceAddress];
            var matrix = ReadMatrix(sourceRange);
            if (matrix.GetLength(0) < 2) throw new InvalidOperationException("Power Query 源区域没有可加载的数据行。");

            var headers = Enumerable.Range(0, matrix.GetLength(1))
                .Select(column => Convert.ToString(matrix[0, column]) ?? "Column" + (column + 1))
                .ToArray();
            if (metadata.Renames != null)
                foreach (var rename in metadata.Renames)
                    for (var column = 0; column < headers.Length; column++)
                        if (string.Equals(headers[column], rename.From, StringComparison.OrdinalIgnoreCase))
                            headers[column] = rename.To;

            var rows = new List<object[]>();
            for (var row = 1; row < matrix.GetLength(0); row++)
            {
                var values = new object[headers.Length];
                var blank = true;
                for (var column = 0; column < headers.Length; column++)
                {
                    var value = matrix[row, column];
                    if (metadata.TrimText && value is string text) value = text.Trim();
                    values[column] = value;
                    if (value != null && !string.IsNullOrWhiteSpace(Convert.ToString(value))) blank = false;
                }
                if (metadata.RemoveBlankRows && blank) continue;
                typeCoercionFailures += ApplyTypes(headers, values, metadata.ColumnTypes);
                rows.Add(values);
            }
            if (metadata.RemoveDuplicates)
                rows = rows.GroupBy(row => string.Join("\u001f", row.Select(NormalizeKey)), StringComparer.Ordinal)
                    .Select(group => group.First()).ToList();

            var selectedIndexes = metadata.SelectColumns != null && metadata.SelectColumns.Length > 0
                ? metadata.SelectColumns.Select(name => Array.FindIndex(headers,
                    header => string.Equals(header, name, StringComparison.OrdinalIgnoreCase)))
                    .Where(index => index >= 0).ToArray()
                : Enumerable.Range(0, headers.Length).ToArray();
            if (selectedIndexes.Length == 0) throw new InvalidOperationException("Power Query 选列规则没有匹配到任何字段。");

            var output = new object[rows.Count + 1, selectedIndexes.Length];
            for (var column = 0; column < selectedIndexes.Length; column++)
                output[0, column] = headers[selectedIndexes[column]];
            for (var row = 0; row < rows.Count; row++)
                for (var column = 0; column < selectedIndexes.Length; column++)
                    output[row + 1, column] = rows[row][selectedIndexes[column]];

            dynamic outputRange = destination.Resize[rows.Count + 1, selectedIndexes.Length];
            outputRange.Value2 = output;
            dynamic listObjects = destinationSheet.ListObjects;
            dynamic listObject = listObjects.Add(
                XlListObjectSourceType.xlSrcRange, outputRange, Type.Missing,
                XlYesNoGuess.xlYes, Type.Missing, Type.Missing);
            listObject.Name = MakeSafeTableName(queryName);
            listObject.TableStyle = "TableStyleMedium4";
            return listObject;
        }

        private static object[,] ReadMatrix(Range range)
        {
            if (range.Rows.Count == 1 && range.Columns.Count == 1)
                return new[,] { { range.Value2 } };
            var source = (object[,])range.Value2;
            var result = new object[source.GetLength(0), source.GetLength(1)];
            for (var row = 1; row <= source.GetLength(0); row++)
                for (var column = 1; column <= source.GetLength(1); column++)
                    result[row - 1, column - 1] = source[row, column];
            return result;
        }

        /// <summary>
        /// 按规则做类型转换;无法解析的单元格保留原值并计入返回的失败数,
        /// 而不是让一个脏值(如整数列里的 "N/A")废掉整个兼容加载。
        /// 源数据经 Value2 读入:数值/日期单元格本来就是 double(日期为 OADate 序列号),
        /// 必须先按数值分流 —— 否则真实日期列会被文本解析判为失败,产生成百上千条假告警。
        /// </summary>
        private static int ApplyTypes(string[] headers, object[] values, IEnumerable<QueryColumnType> types)
        {
            if (types == null) return 0;
            var failures = 0;
            foreach (var rule in types)
            {
                var index = Array.FindIndex(headers,
                    header => string.Equals(header, rule.Field, StringComparison.OrdinalIgnoreCase));
                if (index < 0 || values[index] == null || string.IsNullOrWhiteSpace(Convert.ToString(values[index]))) continue;

                var raw = values[index];
                var type = (rule.Type ?? string.Empty).ToLowerInvariant();

                // ---- 数值单元格(double)分流:天然满足数值类规则;日期规则用 FromOADate ----
                var numericCell = raw as double?;
                if (numericCell.HasValue)
                {
                    switch (type)
                    {
                        case "integer":
                        case "number":
                        case "currency":
                        case "percentage":
                            values[index] = raw;   // 已是数值,无需转换
                            break;
                        case "date":
                        case "datetime":
                            DateTime fromOa;
                            try { fromOa = DateTime.FromOADate(numericCell.Value); }
                            catch { failures++; break; }   // 超出 OADate 有效范围
                            values[index] = type == "date" ? fromOa.Date : fromOa;
                            break;
                        case "text":
                            values[index] = Convert.ToString(raw, CultureInfo.InvariantCulture);
                            break;
                        case "logical":
                            // Excel 里 TRUE/FALSE 经 Value2 也是数值(0/-1),按 0/非0 解释
                            values[index] = numericCell.Value != 0d;
                            break;
                    }
                    continue;
                }

                // ---- 文本单元格:按规则解析 ----
                var text = Convert.ToString(raw).Trim();
                switch (type)
                {
                    case "text":
                        values[index] = text;
                        break;
                    case "integer":
                        long integerValue;
                        if (long.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out integerValue))
                            values[index] = integerValue;
                        else failures++;
                        break;
                    case "number":
                    case "currency":
                    case "percentage":
                        double numericValue;
                        if (double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out numericValue))
                            values[index] = numericValue;
                        else failures++;
                        break;
                    case "date":
                    case "datetime":
                        DateTime dateTimeValue;
                        if (TryParseDateTime(text, out dateTimeValue))
                            values[index] = type == "date" ? dateTimeValue.Date : dateTimeValue;
                        else failures++;
                        break;
                    case "logical":
                        bool logicalValue;
                        if (bool.TryParse(text, out logicalValue))
                            values[index] = logicalValue;
                        else failures++;
                        break;
                }
            }
            return failures;
        }

        private static bool TryParseDateTime(string text, out DateTime value)
        {
            var zhCn = CultureInfo.GetCultureInfo("zh-CN");
            return DateTime.TryParse(text, zhCn, DateTimeStyles.None, out value) ||
                   DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.None, out value);
        }

        private static string NormalizeKey(object value) => value == null ? "<null>" :
            Convert.ToString(value, CultureInfo.InvariantCulture);

        private static void DeleteAllListObjects(Worksheet sheet)
        {
            dynamic listObjects = sheet.ListObjects;
            while (Convert.ToInt32(listObjects.Count) > 0)
                listObjects.Item(1).Delete();
        }

        private static void DeleteFailedQueryConnections(Workbook workbook, string queryName)
        {
            dynamic workbookDynamic = workbook;
            dynamic connections = workbookDynamic.Connections;
            for (var index = Convert.ToInt32(connections.Count); index >= 1; index--)
            {
                dynamic connection = connections.Item(index);
                try
                {
                    var text = Convert.ToString(connection.OLEDBConnection.Connection) ?? string.Empty;
                    if (text.IndexOf("Location=" + queryName, StringComparison.OrdinalIgnoreCase) >= 0)
                        connection.Delete();
                }
                catch { }
            }
        }

        private static void EnsureQueryConnection(Workbook workbook, string queryName, string connectionString)
        {
            dynamic workbookDynamic = workbook;
            dynamic connections = workbookDynamic.Connections;
            var connectionName = "Query - " + queryName;
            for (var index = 1; index <= Convert.ToInt32(connections.Count); index++)
            {
                dynamic connection = connections.Item(index);
                if (string.Equals(Convert.ToString(connection.Name), connectionName, StringComparison.OrdinalIgnoreCase))
                    return;
            }
            connections.Add2(
                connectionName,
                "Power Query 连接 - " + queryName,
                connectionString,
                "\"" + queryName.Replace("\"", "\"\"") + "\"",
                6,
                true,
                false);
        }

        internal static void StyleQueryResult(dynamic resultRange)
        {
            dynamic header = resultRange.Rows.Item(1);
            header.Font.Bold = true;
            header.Font.Color = ColorTranslator.ToOle(Color.White);
            header.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(23, 115, 74));
            header.RowHeight = 24;
            resultRange.Borders.LineStyle = XlLineStyle.xlContinuous;
            resultRange.Borders.Color = ColorTranslator.ToOle(Color.FromArgb(220, 231, 225));
            resultRange.Borders.Weight = XlBorderWeight.xlThin;
            resultRange.VerticalAlignment = XlVAlign.xlVAlignCenter;
            try { resultRange.AutoFilter(); } catch { }
        }

        internal static string MakeSafeTableName(string value)
        {
            var builder = new StringBuilder(MakeSafeTablePrefix(value));
            builder.Append(Math.Abs(DateTime.Now.Ticks % 1000000).ToString(CultureInfo.InvariantCulture));
            return builder.ToString();
        }

        private static string MakeSafeTablePrefix(string value)
        {
            var builder = new StringBuilder("AgentPQ_");
            foreach (var character in value ?? string.Empty)
                builder.Append(char.IsLetterOrDigit(character) || character == '_' ? character : '_');
            builder.Append('_');
            return builder.ToString();
        }

        internal static string EscapeM(string value) => (value ?? string.Empty).Replace("\"", "\"\"");

        private static string MapMType(string value)
        {
            switch ((value ?? string.Empty).Trim().ToLowerInvariant())
            {
                case "text": return "type text";
                case "number": return "type number";
                case "integer": return "Int64.Type";
                case "date": return "type date";
                case "datetime": return "type datetime";
                case "logical": return "type logical";
                case "currency": return "Currency.Type";
                case "percentage": return "Percentage.Type";
                default: throw new ArgumentException("不支持的 Power Query 字段类型：" + value);
            }
        }
    }

    public sealed class ListQueriesOp : IOperation
    {
        public const string PayloadPrefix = "__AGENT_PQ_LIST__";
        public string ToolName => "pq_list_queries";
        public bool IsWriteOperation => false;
        public string Describe() => "列出当前工作簿中的 Power Query 查询";

        public string Execute(AppContext context)
        {
            var workbook = PowerQuerySupport.GetWorkbook(context);
            dynamic workbookDynamic = workbook;
            dynamic queries = workbookDynamic.Queries;
            var result = new List<object>();
            for (var index = 1; index <= Convert.ToInt32(queries.Count); index++)
            {
                dynamic query = queries.Item(index);
                var formula = Convert.ToString(query.Formula) ?? string.Empty;
                result.Add(new
                {
                    name = Convert.ToString(query.Name),
                    description = PowerQuerySupport.GetFriendlyDescription(Convert.ToString(query.Description)),
                    formula_preview = formula.Length <= 300 ? formula : formula.Substring(0, 300) + "…"
                });
            }
            return PayloadPrefix + JsonSerializer.Serialize(new { count = result.Count, queries = result });
        }

        public sealed class Factory : IOperationFactory
        {
            public string ToolName => "pq_list_queries";
            public IOperation Parse(string argumentsJson) => new ListQueriesOp();
        }
    }

    public sealed class CreateRangeQueryOp : IOperation
    {
        private readonly string _sourceSheet;
        private readonly string _sourceAddress;
        private readonly string _queryName;
        private readonly bool _removeBlankRows;
        private readonly bool _trimText;
        private readonly bool _removeDuplicates;
        private readonly bool _replaceExisting;
        private readonly QueryRename[] _renames;
        private readonly QueryColumnType[] _columnTypes;
        private readonly string[] _selectColumns;

        private CreateRangeQueryOp(string sourceSheet, string sourceAddress, string queryName,
            bool removeBlankRows, bool trimText, bool removeDuplicates, bool replaceExisting,
            QueryRename[] renames, QueryColumnType[] columnTypes, string[] selectColumns)
        {
            _sourceSheet = sourceSheet;
            _sourceAddress = sourceAddress;
            _queryName = queryName;
            _removeBlankRows = removeBlankRows;
            _trimText = trimText;
            _removeDuplicates = removeDuplicates;
            _replaceExisting = replaceExisting;
            _renames = renames;
            _columnTypes = columnTypes;
            _selectColumns = selectColumns;
        }

        public string ToolName => "pq_create_from_range";
        public bool IsWriteOperation => true;
        public string Describe() => "根据 " + (_sourceSheet ?? "活动工作表") + "!" + _sourceAddress +
                                    " 创建 Power Query 查询「" + _queryName + "」；源区域保持不变";

        public string Execute(AppContext context)
        {
            var workbook = PowerQuerySupport.GetWorkbook(context);
            var sheet = Cell.CellOperationSupport.GetWorksheet(context, _sourceSheet);
            var range = Cell.CellOperationSupport.GetRange(sheet, _sourceAddress);
            if (range.Rows.Count < 2 || range.Columns.Count < 1)
                throw new ArgumentException("Power Query 源区域必须包含标题行和至少一行数据。");

            var sourceName = "AgentPQSource_" + Math.Abs(DateTime.Now.Ticks % 100000000)
                .ToString(CultureInfo.InvariantCulture);
            var refersTo = "='" + sheet.Name.Replace("'", "''") + "'!" + range.Address;
            dynamic existing = PowerQuerySupport.FindQuery(workbook, _queryName);
            if (existing != null && !_replaceExisting)
                throw new InvalidOperationException("查询「" + _queryName + "」已经存在。请更换名称或设置 replace_existing=true。");

            dynamic sourceWorkbookName = null;
            try
            {
                // Excel.CurrentWorkbook 在部分 Office 构建中不会向 Power Query 暴露隐藏名称。
                // 名称仅作为稳定的数据源引用，不改动源区域内容或格式。
                sourceWorkbookName = workbook.Names.Add(sourceName, refersTo, true);
                var formula = PowerQuerySupport.BuildRangeQueryFormula(
                    sourceName, _removeBlankRows, _trimText, _removeDuplicates,
                    _renames, _columnTypes, _selectColumns);
                var metadata = new AgentPowerQueryMetadata
                {
                    SourceSheet = sheet.Name,
                    SourceAddress = Convert.ToString(range.Address),
                    SourceName = sourceName,
                    RemoveBlankRows = _removeBlankRows,
                    TrimText = _trimText,
                    RemoveDuplicates = _removeDuplicates,
                    Renames = _renames ?? new QueryRename[0],
                    ColumnTypes = _columnTypes ?? new QueryColumnType[0],
                    SelectColumns = _selectColumns ?? new string[0]
                };
                var description = PowerQuerySupport.BuildDescription(
                    "由 Agent for Excel 创建；源区域 " + sheet.Name + "!" + range.Address, metadata);
                if (existing == null)
                {
                    dynamic workbookDynamic = workbook;
                    workbookDynamic.Queries.Add(
                        _queryName,
                        formula,
                        description);
                }
                else
                {
                    existing.Formula = formula;
                    existing.Description = description;
                }
                return "已创建 Power Query 查询「" + _queryName + "」，包含标题提升" +
                       (_removeBlankRows ? "、删除空行" : string.Empty) +
                       (_trimText ? "、文本清理" : string.Empty) +
                       (_removeDuplicates ? "、删除重复项" : string.Empty) +
                       "；源数据未修改。";
            }
            catch
            {
                if (existing == null && sourceWorkbookName != null)
                    try { sourceWorkbookName.Delete(); } catch { }
                throw;
            }
        }

        public sealed class Factory : IOperationFactory
        {
            public string ToolName => "pq_create_from_range";
            public IOperation Parse(string argumentsJson)
            {
                using (var document = JsonDocument.Parse(argumentsJson))
                {
                    var root = document.RootElement;
                    var address = ReadString(root, "source_address");
                    var name = ReadString(root, "query_name");
                    if (string.IsNullOrWhiteSpace(address)) throw new ArgumentException("source_address 不能为空。");
                    if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("query_name 不能为空。");
                    return new CreateRangeQueryOp(
                        ReadString(root, "source_sheet"), address, name,
                        ReadBool(root, "remove_blank_rows", true),
                        ReadBool(root, "trim_text", true),
                        ReadBool(root, "remove_duplicates", true),
                        ReadBool(root, "replace_existing", false),
                        ReadRenames(root), ReadTypes(root), ReadStringArray(root, "select_columns"));
                }
            }

            private static string ReadString(JsonElement root, string name) =>
                root.TryGetProperty(name, out var value) && value.ValueKind == JsonValueKind.String
                    ? value.GetString()?.Trim() : null;

            private static bool ReadBool(JsonElement root, string name, bool fallback) =>
                root.TryGetProperty(name, out var value) && (value.ValueKind == JsonValueKind.True || value.ValueKind == JsonValueKind.False)
                    ? value.GetBoolean() : fallback;

            private static string[] ReadStringArray(JsonElement root, string name)
            {
                if (!root.TryGetProperty(name, out var value)) return new string[0];
                if (value.ValueKind != JsonValueKind.Array) throw new ArgumentException(name + " 必须是字符串数组。");
                return value.EnumerateArray().Select(item => item.GetString()?.Trim())
                    .Where(item => !string.IsNullOrWhiteSpace(item)).ToArray();
            }

            private static QueryRename[] ReadRenames(JsonElement root)
            {
                if (!root.TryGetProperty("rename_columns", out var value)) return new QueryRename[0];
                var result = new List<QueryRename>();
                foreach (var item in value.EnumerateArray())
                    result.Add(new QueryRename { From = ReadString(item, "from"), To = ReadString(item, "to") });
                if (result.Any(item => string.IsNullOrWhiteSpace(item.From) || string.IsNullOrWhiteSpace(item.To)))
                    throw new ArgumentException("rename_columns 的 from/to 不能为空。");
                return result.ToArray();
            }

            private static QueryColumnType[] ReadTypes(JsonElement root)
            {
                if (!root.TryGetProperty("column_types", out var value)) return new QueryColumnType[0];
                var result = new List<QueryColumnType>();
                foreach (var item in value.EnumerateArray())
                    result.Add(new QueryColumnType { Field = ReadString(item, "field"), Type = ReadString(item, "type") });
                if (result.Any(item => string.IsNullOrWhiteSpace(item.Field) || string.IsNullOrWhiteSpace(item.Type)))
                    throw new ArgumentException("column_types 的 field/type 不能为空。");
                return result.ToArray();
            }
        }
    }

    public sealed class LoadQueryOp : IOperation
    {
        private readonly string _queryName;
        private readonly string _destinationSheet;
        private readonly string _destinationAddress;

        private LoadQueryOp(string queryName, string destinationSheet, string destinationAddress)
        {
            _queryName = queryName;
            _destinationSheet = destinationSheet;
            _destinationAddress = destinationAddress;
        }

        public string ToolName => "pq_load_to_sheet";
        public bool IsWriteOperation => true;
        public string Describe() => "把 Power Query 查询「" + _queryName + "」加载到新工作表「" + _destinationSheet + "」";

        public string Execute(AppContext context)
        {
            var workbook = PowerQuerySupport.GetWorkbook(context);
            if (PowerQuerySupport.FindQuery(workbook, _queryName) == null)
                throw new ArgumentException("找不到 Power Query 查询「" + _queryName + "」。");
            Worksheet sheet = null;
            try
            {
                sheet = Analysis.AnalysisSheetSupport.CreateUniqueWorksheet(context, _destinationSheet);
                var destination = sheet.Range[_destinationAddress];
                var outcome = PowerQuerySupport.LoadQueryToSheet(
                    workbook, sheet, _queryName, destination, _queryName);
                dynamic listObject = outcome.ListObject;
                dynamic resultRange = listObject.Range;
                var rows = Convert.ToInt32(resultRange.Rows.Count);
                var columns = Convert.ToInt32(resultRange.Columns.Count);
                PowerQuerySupport.StyleQueryResult(resultRange);
                sheet.Columns.AutoFit();
                sheet.Activate();
                var coercionNote = outcome.TypeCoercionFailures > 0
                    ? $"（{outcome.TypeCoercionFailures} 个单元格无法按指定类型转换，已保留原值）"
                    : string.Empty;
                return "已将查询「" + _queryName + "」加载到 " + sheet.Name + "!" +
                       Convert.ToString(resultRange.Address) + "，结果为 " + Math.Max(0, rows - 1) +
                       " 行 × " + columns + " 列" +
                       (outcome.UsedCompatibilityEngine ? "（已自动使用兼容清洗引擎）" : string.Empty) +
                       coercionNote + "。";
            }
            catch
            {
                if (sheet != null) Analysis.AnalysisSheetSupport.DeleteWorksheetSilently(context, sheet);
                throw;
            }
        }

        public sealed class Factory : IOperationFactory
        {
            public string ToolName => "pq_load_to_sheet";
            public IOperation Parse(string argumentsJson)
            {
                using (var document = JsonDocument.Parse(argumentsJson))
                {
                    var root = document.RootElement;
                    var queryName = ReadString(root, "query_name");
                    if (string.IsNullOrWhiteSpace(queryName)) throw new ArgumentException("query_name 不能为空。");
                    return new LoadQueryOp(queryName,
                        ReadString(root, "destination_sheet") ?? "Agent查询结果",
                        ReadString(root, "destination_address") ?? "A1");
                }
            }
            private static string ReadString(JsonElement root, string name) =>
                root.TryGetProperty(name, out var value) && value.ValueKind == JsonValueKind.String
                    ? value.GetString()?.Trim() : null;
        }
    }

    public sealed class RefreshQueryOp : IOperation
    {
        private readonly string _queryName;
        private RefreshQueryOp(string queryName) { _queryName = queryName; }
        public string ToolName => "pq_refresh";
        public bool IsWriteOperation => true;
        public string Describe() => "刷新 Power Query 查询「" + _queryName + "」及其加载结果";

        public string Execute(AppContext context)
        {
            var workbook = PowerQuerySupport.GetWorkbook(context);
            if (PowerQuerySupport.FindQuery(workbook, _queryName) == null)
                throw new ArgumentException("找不到 Power Query 查询「" + _queryName + "」。");
            var refreshed = 0;
            var coercionFailures = 0;
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                dynamic listObjects = sheet.ListObjects;
                for (var index = 1; index <= Convert.ToInt32(listObjects.Count); index++)
                {
                    dynamic listObject = listObjects.Item(index);
                    if (!PowerQuerySupport.IsFallbackTableForQuery(listObject, _queryName)) continue;
                    try
                    {
                        dynamic queryTable = listObject.QueryTable;
                        queryTable.BackgroundQuery = false;
                        queryTable.Refresh(false);
                        refreshed++;
                    }
                    catch
                    {
                        int failures;
                        if (PowerQuerySupport.RefreshFallbackTable(workbook, sheet, listObject, _queryName, out failures) != null)
                        {
                            refreshed++;
                            coercionFailures += failures;
                        }
                    }
                    break;
                }
            }
            if (refreshed == 0) return "查询「" + _queryName + "」尚未加载到工作表；查询定义已保留。";
            var refreshNote = coercionFailures > 0
                ? $"（{coercionFailures} 个单元格无法按指定类型转换，已保留原值）"
                : string.Empty;
            return "已刷新查询「" + _queryName + "」的 " + refreshed + " 个加载结果" + refreshNote + "。";
        }

        public sealed class Factory : IOperationFactory
        {
            public string ToolName => "pq_refresh";
            public IOperation Parse(string argumentsJson)
            {
                using (var document = JsonDocument.Parse(argumentsJson))
                {
                    if (!document.RootElement.TryGetProperty("query_name", out var value) || value.ValueKind != JsonValueKind.String)
                        throw new ArgumentException("query_name 不能为空。");
                    return new RefreshQueryOp(value.GetString().Trim());
                }
            }
        }
    }
}
