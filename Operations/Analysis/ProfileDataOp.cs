using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;

namespace AgentForExcel.Operations.Analysis
{
    /// <summary>在任何分析或制图前，对二维数据做只读字段画像和高价值质量检查。</summary>
    public sealed class ProfileDataOp : IOperation
    {
        public const string PayloadPrefix = "__AGENT_DATA_PROFILE__";
        public string ToolName => "data_profile";
        public bool IsWriteOperation => false;

        private readonly string _sheetName;
        private readonly string _address;

        private ProfileDataOp(string sheetName, string address)
        {
            _sheetName = sheetName;
            _address = address;
        }

        public string Describe()
        {
            var target = string.IsNullOrWhiteSpace(_sheetName) ? "活动工作表" : "工作表「" + _sheetName + "」";
            return $"体检 {target}!{_address} 的字段类型、完整性、基数、异常值和分析粒度";
        }

        public string Execute(AppContext context)
        {
            var sheet = Cell.CellOperationSupport.GetWorksheet(context, _sheetName);
            var range = Cell.CellOperationSupport.GetRange(sheet, _address);
            var rowCount = Convert.ToInt32(range.Rows.Count);
            var columnCount = Convert.ToInt32(range.Columns.Count);
            if (rowCount < 2 || columnCount < 1)
                throw new ArgumentException("数据体检区域至少需要标题行和一行数据。");
            if (columnCount > 100 || (long)rowCount * columnCount > 500000)
                throw new ArgumentException("单次数据体检最多支持 100 列或 500,000 个单元格，请缩小范围或分批检查。");

            var headers = ValidateHeaders(range, columnCount);
            var fields = new List<FieldProfile>();
            var warnings = new List<string>();
            var derivedDimensions = new List<string>();
            var rowSignatures = new Dictionary<string, int>(StringComparer.Ordinal);

            for (var column = 1; column <= columnCount; column++)
            {
                var profile = ProfileColumn(range, column, headers[column - 1], rowCount - 1);
                fields.Add(profile);
                AddFieldWarnings(profile, warnings);
                if (profile.InferredType == "date")
                    AddDerivedDimensions(profile.Name, derivedDimensions);
            }

            for (var row = 2; row <= rowCount; row++)
            {
                var signature = BuildRowSignature(range, row, columnCount);
                int count;
                rowSignatures.TryGetValue(signature, out count);
                rowSignatures[signature] = count + 1;
            }

            var duplicateRows = 0;
            foreach (var pair in rowSignatures)
                if (pair.Value > 1) duplicateRows += pair.Value - 1;
            if (duplicateRows > 0)
                warnings.Add($"发现 {duplicateRows} 条完全重复记录；分析前需要确认预期粒度或去重规则。");

            var dimensionCandidates = new List<string>();
            var measureCandidates = new List<string>();
            var timeCandidates = new List<string>();
            foreach (var field in fields)
            {
                if (field.Role == "dimension") dimensionCandidates.Add(field.Name + "(" + field.DistinctCount + ")");
                if (field.Role == "measure") measureCandidates.Add(field.Name);
                if (field.Role == "time") timeCandidates.Add(field.Name);
            }

            var payload = new
            {
                kind = "data_profile",
                sheet = Convert.ToString(sheet.Name),
                address = range.Address,
                data_rows = rowCount - 1,
                columns = columnCount,
                duplicate_rows = duplicateRows,
                duplicate_rate = Math.Round((double)duplicateRows / Math.Max(1, rowCount - 1), 4),
                fields,
                dimension_candidates = dimensionCandidates,
                measure_candidates = measureCandidates,
                time_candidates = timeCandidates,
                suggested_derived_dimensions = derivedDimensions,
                warnings,
                chart_guardrails = new[]
                {
                    "同一分类出现多行时，先确认粒度并按 sum/average/count 等规则聚合，不得把重复分类直接画到横轴。",
                    "饼图/环形图建议不超过 8 个分类；排名图建议展示 Top 5-12，其余合并为其他。",
                    "折线图至少需要 8 个有序时间点；密集时间序列保留数据点但减少横轴标签，不逐点显示数值。",
                    "高基数字段不得直接作为分类轴，应先派生时间粒度、分箱、Top-N 或业务分组。"
                }
            };
            return PayloadPrefix + JsonSerializer.Serialize(payload);
        }

        private static List<string> ValidateHeaders(Range range, int columnCount)
        {
            var headers = new List<string>();
            var unique = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (var column = 1; column <= columnCount; column++)
            {
                var header = Convert.ToString(((Range)range.Cells[1, column]).Value2)?.Trim();
                if (string.IsNullOrWhiteSpace(header)) throw new ArgumentException("数据体检区域存在空白字段名。");
                if (!unique.Add(header)) throw new ArgumentException("数据体检区域存在重复字段名：「" + header + "」。");
                headers.Add(header);
            }
            return headers;
        }

        private static FieldProfile ProfileColumn(Range range, int column, string header, int totalRows)
        {
            var profile = new FieldProfile { Name = header, TotalRows = totalRows };
            var distinct = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var numericValues = new List<double>();
            var examples = new List<string>();
            var dateLike = 0;
            var booleanLike = 0;
            var numericLike = 0;

            for (var row = 2; row <= totalRows + 1; row++)
            {
                var value = ((Range)range.Cells[row, column]).Value2;
                var text = Normalize(value);
                if (string.IsNullOrWhiteSpace(text))
                {
                    profile.MissingCount++;
                    continue;
                }

                profile.NonBlankCount++;
                distinct.Add(text);
                if (examples.Count < 3 && !examples.Contains(text)) examples.Add(Truncate(text, 40));

                double number;
                if (TryNumber(value, out number))
                {
                    numericLike++;
                    numericValues.Add(number);
                    if (Math.Abs(number) < 0.0000001) profile.ZeroCount++;
                }
                DateTime ignoredDate;
                if (TryTemporal(value, text, header, out ignoredDate)) dateLike++;
                bool ignoredBoolean;
                if (bool.TryParse(text, out ignoredBoolean) || text == "是" || text == "否") booleanLike++;
            }

            profile.DistinctCount = distinct.Count;
            profile.MissingRate = Math.Round((double)profile.MissingCount / Math.Max(1, totalRows), 4);
            profile.DistinctRate = Math.Round((double)profile.DistinctCount / Math.Max(1, profile.NonBlankCount), 4);
            profile.AverageRowsPerValue = Math.Round((double)profile.NonBlankCount / Math.Max(1, profile.DistinctCount), 2);
            profile.ZeroRate = Math.Round((double)profile.ZeroCount / Math.Max(1, profile.NonBlankCount), 4);
            profile.Examples = examples;

            var threshold = Math.Max(1, (int)Math.Ceiling(profile.NonBlankCount * 0.9));
            if (profile.NonBlankCount == 0) profile.InferredType = "empty";
            else if (dateLike >= threshold) profile.InferredType = "date";
            else if (booleanLike >= threshold) profile.InferredType = "boolean";
            else if (numericLike >= threshold) profile.InferredType = "number";
            else if (numericLike > 0 || dateLike > 0) profile.InferredType = "mixed";
            else profile.InferredType = "text";

            if (profile.InferredType == "date") profile.Role = "time";
            else if (profile.InferredType == "number") profile.Role = "measure";
            else if (LooksLikeId(header) && profile.DistinctRate > 0.95) profile.Role = "identifier";
            else if (profile.DistinctCount <= Math.Min(100, Math.Max(8, profile.NonBlankCount / 2))) profile.Role = "dimension";
            else profile.Role = "text";

            if (numericValues.Count > 0 && profile.InferredType != "date")
            {
                numericValues.Sort();
                profile.Min = numericValues[0];
                profile.Max = numericValues[numericValues.Count - 1];
                double sum = 0;
                foreach (var value in numericValues) sum += value;
                profile.Mean = Math.Round(sum / numericValues.Count, 4);
                profile.OutlierCount = CountIqrOutliers(numericValues);
            }
            return profile;
        }

        private static void AddFieldWarnings(FieldProfile profile, IList<string> warnings)
        {
            if (profile.MissingRate >= 0.1)
                warnings.Add($"字段「{profile.Name}」缺失率为 {profile.MissingRate:P1}，可能影响分组或计算。" );
            if (profile.InferredType == "mixed")
                warnings.Add($"字段「{profile.Name}」存在混合类型，需要统一格式后再分析。" );
            if (profile.OutlierCount > 0)
                warnings.Add($"字段「{profile.Name}」按 IQR 规则发现 {profile.OutlierCount} 个潜在异常值。" );
            if (profile.ZeroRate >= 0.3 && profile.InferredType == "number")
                warnings.Add($"字段「{profile.Name}」零值占比为 {profile.ZeroRate:P1}，需要确认零值是否代表真实业务含义。" );
            if ((profile.Role == "text" || profile.Role == "identifier") && profile.DistinctCount > 50)
                warnings.Add($"字段「{profile.Name}」有 {profile.DistinctCount} 个不同值，不适合直接作为图表分类轴。" );
            if ((profile.Role == "time" || profile.Role == "dimension") && profile.AverageRowsPerValue > 1.05)
                warnings.Add($"字段「{profile.Name}」每个值平均对应 {profile.AverageRowsPerValue:0.##} 行明细；作为横轴前必须先确认 sum/average/count 等聚合口径。" );
        }

        private static void AddDerivedDimensions(string fieldName, IList<string> dimensions)
        {
            var normalized = (fieldName ?? string.Empty).ToLowerInvariant();
            dimensions.Add(fieldName + " → 年");
            if (normalized.Contains("季度") || normalized.Contains("quarter")) return;
            dimensions.Add(fieldName + " → 季度");
            if (normalized.Contains("月份") || normalized.Contains("年月") || normalized.Contains("month")) return;
            dimensions.Add(fieldName + " → 年月");
            dimensions.Add(fieldName + " → 星期");
        }

        private static int CountIqrOutliers(IList<double> sorted)
        {
            if (sorted.Count < 8) return 0;
            var q1 = sorted[(int)Math.Floor((sorted.Count - 1) * 0.25)];
            var q3 = sorted[(int)Math.Floor((sorted.Count - 1) * 0.75)];
            var iqr = q3 - q1;
            if (Math.Abs(iqr) < 0.0000001) return 0;
            var lower = q1 - 1.5 * iqr;
            var upper = q3 + 1.5 * iqr;
            var count = 0;
            foreach (var value in sorted) if (value < lower || value > upper) count++;
            return count;
        }

        private static string BuildRowSignature(Range range, int row, int columnCount)
        {
            var builder = new StringBuilder();
            for (var column = 1; column <= columnCount; column++)
            {
                if (column > 1) builder.Append('\u001F');
                builder.Append(Normalize(((Range)range.Cells[row, column]).Value2).ToLowerInvariant());
            }
            return builder.ToString();
        }

        private static bool TryTemporal(object value, string text, string header, out DateTime date)
        {
            if (value is DateTime)
            {
                date = (DateTime)value;
                return true;
            }
            if (Regex.IsMatch(text, @"^\d{4}[-/]?Q[1-4]$", RegexOptions.IgnoreCase) ||
                Regex.IsMatch(text, @"^\d{4}年第?[一二三四1234]季度$"))
            {
                date = new DateTime(Convert.ToInt32(text.Substring(0, 4)), 1, 1);
                return true;
            }
            if (LooksLikeDateHeader(header) && DateTime.TryParse(text, CultureInfo.CurrentCulture, DateTimeStyles.None, out date))
                return true;
            double serial;
            if (LooksLikeDateHeader(header) && TryNumber(value, out serial) && serial > 20000 && serial < 100000)
            {
                try { date = DateTime.FromOADate(serial); return true; } catch { }
            }
            date = default(DateTime);
            return false;
        }

        private static bool LooksLikeDateHeader(string header)
        {
            var value = (header ?? string.Empty).ToLowerInvariant();
            return value.Contains("日期") || value.Contains("时间") || value.Contains("季度") ||
                   value.Contains("月份") || value.Contains("年月") || value == "年" ||
                   value.Contains("date") || value.Contains("time") || value.Contains("quarter") || value.Contains("month");
        }

        private static bool LooksLikeId(string header)
        {
            var value = (header ?? string.Empty).ToLowerInvariant();
            return value == "id" || value.EndsWith("id") || value.Contains("编号") || value.Contains("编码") || value.Contains("序号");
        }

        private static bool TryNumber(object value, out double number)
        {
            if (value == null) { number = 0; return false; }
            try
            {
                number = Convert.ToDouble(value, CultureInfo.InvariantCulture);
                return !double.IsNaN(number) && !double.IsInfinity(number);
            }
            catch { number = 0; return false; }
        }

        private static string Normalize(object value)
        {
            return (Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty).Trim();
        }

        private static string Truncate(string value, int length)
        {
            return value.Length <= length ? value : value.Substring(0, length - 1) + "…";
        }

        private sealed class FieldProfile
        {
            public string Name { get; set; }
            public string InferredType { get; set; }
            public string Role { get; set; }
            public int TotalRows { get; set; }
            public int NonBlankCount { get; set; }
            public int MissingCount { get; set; }
            public double MissingRate { get; set; }
            public int DistinctCount { get; set; }
            public double DistinctRate { get; set; }
            public double AverageRowsPerValue { get; set; }
            public int ZeroCount { get; set; }
            public double ZeroRate { get; set; }
            public double Min { get; set; }
            public double Max { get; set; }
            public double Mean { get; set; }
            public int OutlierCount { get; set; }
            public List<string> Examples { get; set; }
        }

        public sealed class Factory : IOperationFactory
        {
            public string ToolName => "data_profile";

            public IOperation Parse(string argumentsJson)
            {
                using (var document = JsonDocument.Parse(argumentsJson))
                {
                    var root = document.RootElement;
                    var address = ReadString(root, "address");
                    if (string.IsNullOrWhiteSpace(address)) throw new ArgumentException("address 不能为空。");
                    return new ProfileDataOp(ReadString(root, "sheet"), address);
                }
            }

            private static string ReadString(JsonElement root, string name)
            {
                return root.TryGetProperty(name, out var value) && value.ValueKind == JsonValueKind.String
                    ? value.GetString()?.Trim()
                    : null;
            }
        }
    }
}
