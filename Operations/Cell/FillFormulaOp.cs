using System;
using System.Text.Json;
using Microsoft.Office.Interop.Excel;

namespace AgentForExcel.Operations.Cell
{
    /// <summary>向区域填充 A1 或 R1C1 公式。</summary>
    public sealed class FillFormulaOp : IOperation
    {
        public string ToolName => "cell_fill_formula";
        public bool IsWriteOperation => true;

        private readonly string _sheetName;
        private readonly string _address;
        private readonly string _formula;
        private readonly bool _useR1C1;

        private FillFormulaOp(string sheetName, string address, string formula, bool useR1C1)
        {
            _sheetName = sheetName;
            _address = address;
            _formula = formula;
            _useR1C1 = useR1C1;
        }

        public string Describe()
        {
            var where = string.IsNullOrWhiteSpace(_sheetName) ? "活动工作表" : "工作表「" + _sheetName + "」";
            var notation = _useR1C1 ? "R1C1" : "A1";
            return $"向 {where} 的 {_address} 填充 {notation} 公式：{_formula}";
        }

        public string Execute(AppContext context)
        {
            CellOperationSupport.ValidateFormula(_formula);
            var sheet = CellOperationSupport.GetWorksheet(context, _sheetName);
            var target = CellOperationSupport.GetRange(sheet, _address);

            if (_useR1C1) target.FormulaR1C1 = _formula;
            else target.Formula = _formula;

            var firstCell = (Range)target.Cells[1, 1];
            var writtenFormula = _useR1C1 ? firstCell.FormulaR1C1 : firstCell.Formula;
            return $"已向 {sheet.Name}!{target.Address} 填充公式，共 {target.CountLarge} 个单元格；首个公式为 {writtenFormula}。";
        }

        public sealed class Factory : IOperationFactory
        {
            public string ToolName => "cell_fill_formula";

            public IOperation Parse(string argumentsJson)
            {
                using (var doc = JsonDocument.Parse(argumentsJson))
                {
                    var root = doc.RootElement;
                    var sheet = ReadString(root, "sheet");
                    var address = ReadString(root, "address");
                    var formula = ReadString(root, "formula");
                    var useR1C1 = root.TryGetProperty("use_r1c1", out var notation) &&
                                  (notation.ValueKind == JsonValueKind.True || notation.ValueKind == JsonValueKind.False) &&
                                  notation.GetBoolean();

                    if (string.IsNullOrWhiteSpace(address)) throw new ArgumentException("address 不能为空。");
                    CellOperationSupport.ValidateFormula(formula);
                    return new FillFormulaOp(sheet, address, formula, useR1C1);
                }
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
