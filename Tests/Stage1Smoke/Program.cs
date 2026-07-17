using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.IO;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace AgentForExcel.Stage1Smoke
{
    internal static class Program
    {
        private static Assembly _agentAssembly;
        private static object _context;

        [STAThread]
        private static int Main(string[] args)
        {
            dynamic excel = null;
            dynamic workbook = null;
            dynamic worksheet = null;
            try
            {
                var agentDirectory = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, ".."));
                AppDomain.CurrentDomain.AssemblyResolve += (sender, eventArgs) =>
                {
                    var dependencyName = new AssemblyName(eventArgs.Name).Name + ".dll";
                    var dependencyPath = Path.Combine(agentDirectory, dependencyName);
                    return File.Exists(dependencyPath) ? Assembly.LoadFrom(dependencyPath) : null;
                };

                var agentPath = Path.Combine(agentDirectory, "AgentForExcel.dll");
                _agentAssembly = Assembly.LoadFrom(agentPath);

                if (args != null && Array.IndexOf(args, "--ui-only") >= 0)
                {
                    if (System.Windows.Application.Current == null)
                        new System.Windows.Application { ShutdownMode = System.Windows.ShutdownMode.OnExplicitShutdown };
                    ExportUiPreviewsIfRequested();
                    Console.WriteLine("PASS");
                    return 0;
                }
                if (args != null && Array.IndexOf(args, "--settings-only") >= 0)
                    return RunSettingsAndCatalogSmoke();
                var permissionOnly = args != null && Array.IndexOf(args, "--permission-only") >= 0;

                excel = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application", true));
                excel.AutomationSecurity = 1;
                excel.Visible = false;
                excel.DisplayAlerts = false;
                workbook = excel.Workbooks.Add();
                worksheet = workbook.Worksheets.Item(1);
                worksheet.Name = "Stage1Test";

                _context = CreateContext(excel);

                if (permissionOnly)
                {
                    TestSafeAutomationPermissions(worksheet);
                    Console.WriteLine("PASS");
                    Console.WriteLine("实时选区锁定与安全自动化权限验证通过。");
                    return 0;
                }

                var write = InvokeOperation(
                    "AgentForExcel.Operations.Cell.WriteRangeOp+Factory",
                    "{\"sheet\":\"Stage1Test\",\"address\":\"A1\",\"values\":[[\"产品\",\"数量\"],[\"A\",10],[\"B\",20]]}");
                var formula = InvokeOperation(
                    "AgentForExcel.Operations.Cell.FillFormulaOp+Factory",
                    "{\"sheet\":\"Stage1Test\",\"address\":\"C2:C3\",\"formula\":\"=RC[-1]*2\",\"use_r1c1\":true}");
                var read = InvokeOperation(
                    "AgentForExcel.Operations.Cell.ReadRangeOp+Factory",
                    "{\"sheet\":\"Stage1Test\",\"address\":\"A1:C3\"}");
                var format = InvokeOperation(
                    "AgentForExcel.Operations.Cell.FormatRangeOp+Factory",
                    "{\"sheet\":\"Stage1Test\",\"address\":\"A1:C1\",\"bold\":true,\"fill_color\":\"#E9F4EC\",\"font_color\":\"#16764A\",\"horizontal_alignment\":\"center\",\"add_borders\":true,\"autofit_columns\":true}");
                var chart = InvokeOperation(
                    "AgentForExcel.Operations.Chart.CreateChartOp+Factory",
                    "{\"source_sheet\":\"Stage1Test\",\"source_address\":\"A1:C3\",\"destination_sheet\":\"Stage1Test\",\"anchor_address\":\"E5\",\"chart_type\":\"column\",\"title\":\"产品数量与金额\",\"name\":\"SmokeChart\"}");
                var barChart = InvokeOperation(
                    "AgentForExcel.Operations.Chart.CreateChartOp+Factory",
                    "{\"source_sheet\":\"Stage1Test\",\"source_address\":\"A1:C3\",\"destination_sheet\":\"Stage1Test\",\"anchor_address\":\"E22\",\"chart_type\":\"bar\",\"title\":\"横向对比\",\"name\":\"SmokeBar\"}");
                var lineChart = InvokeOperation(
                    "AgentForExcel.Operations.Chart.CreateChartOp+Factory",
                    "{\"source_sheet\":\"Stage1Test\",\"source_address\":\"A1:C3\",\"destination_sheet\":\"Stage1Test\",\"anchor_address\":\"E39\",\"chart_type\":\"line\",\"title\":\"趋势图\",\"name\":\"SmokeLine\"}");
                var areaChart = InvokeOperation(
                    "AgentForExcel.Operations.Chart.CreateChartOp+Factory",
                    "{\"source_sheet\":\"Stage1Test\",\"source_address\":\"A1:C3\",\"destination_sheet\":\"Stage1Test\",\"anchor_address\":\"E56\",\"chart_type\":\"area\",\"title\":\"累计趋势\",\"name\":\"SmokeArea\"}");
                var scatterChart = InvokeOperation(
                    "AgentForExcel.Operations.Chart.CreateChartOp+Factory",
                    "{\"source_sheet\":\"Stage1Test\",\"source_address\":\"B1:C3\",\"destination_sheet\":\"Stage1Test\",\"anchor_address\":\"E73\",\"chart_type\":\"scatter\",\"title\":\"相关关系\",\"name\":\"SmokeScatter\"}");
                InvokeOperation(
                    "AgentForExcel.Operations.Cell.WriteRangeOp+Factory",
                    "{\"sheet\":\"Stage1Test\",\"address\":\"G1\",\"values\":[[\"区域\",\"产品\",\"销售额\"],[\"北区\",\"A\",100],[\"南区\",\"A\",80],[\"北区\",\"B\",120],[\"南区\",\"B\",90],[\"北区\",\"A\",60]]}");
                var analysisView = InvokeOperation(
                    "AgentForExcel.Operations.Analysis.CreateAnalysisViewOp+Factory",
                    "{\"source_sheet\":\"Stage1Test\",\"source_address\":\"G1:I6\",\"analysis_sheet_name\":\"Agent分析\",\"sort_by\":[{\"field\":\"销售额\",\"direction\":\"desc\"}]}");
                var pivot = InvokeOperation(
                    "AgentForExcel.Operations.Pivot.CreatePivotTableOp+Factory",
                    "{\"source_sheet\":\"Stage1Test\",\"source_address\":\"G1:I6\",\"destination_sheet\":\"PivotTest\",\"destination_address\":\"A1\",\"name\":\"SmokePivot\",\"rows\":[\"区域\"],\"columns\":[\"产品\"],\"values\":[{\"field\":\"销售额\",\"function\":\"sum\",\"caption\":\"销售额合计\"}]}" );
                InvokeOperation(
                    "AgentForExcel.Operations.Cell.WriteRangeOp+Factory",
                    "{\"sheet\":\"Stage1Test\",\"address\":\"K1\",\"values\":[[\"Category\",\"Budget\",\"Actual\"],[\"住房\",3000,3000],[\"储蓄\",2000,2000],[\"餐饮\",1500,1820],[\"娱乐\",1000,1450],[\"交通\",800,650],[\"购物\",800,420],[\"医疗\",500,280],[\"Total\",9600,9620]]}");
                var doughnut = InvokeOperation(
                    "AgentForExcel.Operations.Chart.CreateChartOp+Factory",
                    "{\"source_sheet\":\"Stage1Test\",\"source_address\":\"K1:M9\",\"chart_type\":\"doughnut\",\"title\":\"各类别实际消费占比\",\"category_field\":\"Category\",\"value_field\":\"Actual\",\"exclude_categories\":[\"储蓄\"],\"sort_descending\":true,\"show_data_labels\":true,\"show_percentage\":true,\"palette\":\"emerald\"}");
                InvokeOperation(
                    "AgentForExcel.Operations.Cell.WriteRangeOp+Factory",
                    "{\"sheet\":\"Stage1Test\",\"address\":\"U1\",\"values\":[[\"季度\",\"Volume\"],[\"2014-Q1\",41632500],[\"2014-Q1\",45517700],[\"2014-Q2\",68674700],[\"2014-Q2\",53293800],[\"2014-Q3\",69087900],[\"2014-Q3\",45339800]]}");
                var profile = InvokeOperation(
                    "AgentForExcel.Operations.Analysis.ProfileDataOp+Factory",
                    "{\"sheet\":\"Stage1Test\",\"address\":\"U1:V7\"}");
                var aggregatedTrend = InvokeOperation(
                    "AgentForExcel.Operations.Chart.CreateChartOp+Factory",
                    "{\"source_sheet\":\"Stage1Test\",\"source_address\":\"U1:V7\",\"chart_type\":\"line\",\"title\":\"季度成交量\",\"category_field\":\"季度\",\"value_field\":\"Volume\",\"aggregation\":\"auto\",\"show_data_labels\":false,\"max_categories\":12}");
                InvokeOperation(
                    "AgentForExcel.Operations.Cell.WriteRangeOp+Factory",
                    "{\"sheet\":\"Stage1Test\",\"address\":\"O1\",\"values\":[[\"日期\",\"区域\",\"类别\",\"产品\",\"销售额\",\"数量\"],[\"2026-01\",\"华东\",\"办公\",\"A\",12800,12],[\"2026-01\",\"华南\",\"办公\",\"B\",9600,9],[\"2026-01\",\"华东\",\"家居\",\"C\",7200,7],[\"2026-02\",\"华北\",\"办公\",\"A\",15600,14],[\"2026-02\",\"华南\",\"家居\",\"C\",8300,8],[\"2026-02\",\"华东\",\"数码\",\"D\",18600,11],[\"2026-03\",\"华北\",\"数码\",\"D\",20900,13],[\"2026-03\",\"华南\",\"办公\",\"B\",11200,10],[\"2026-03\",\"华东\",\"家居\",\"C\",9900,9],[\"2026-04\",\"华北\",\"办公\",\"A\",17400,15],[\"2026-04\",\"华南\",\"数码\",\"D\",22100,14],[\"2026-04\",\"华东\",\"家居\",\"C\",10600,10]]}");
                var dashboardSourceDateBefore = Convert.ToString(worksheet.Range("O2").Value2);
                var dashboardSourceValueBefore = Convert.ToDouble(worksheet.Range("S2").Value2);
                var dashboard = InvokeOperation(
                    "AgentForExcel.Operations.Dashboard.CreateDashboardOp+Factory",
                    "{\"source_sheet\":\"Stage1Test\",\"source_address\":\"O1:T13\",\"dashboard_sheet_name\":\"Agent看板\",\"title\":\"销售经营看板\",\"date_field\":\"日期\",\"category_field\":\"产品\",\"series_field\":\"区域\",\"value_field\":\"销售额\",\"filter_fields\":[\"区域\",\"类别\"],\"aggregation\":\"sum\",\"top_n\":5,\"number_format\":\"¥#,##0\"}");
                var dropdownDashboard = InvokeOperation(
                    "AgentForExcel.Operations.Dashboard.CreateDashboardOp+Factory",
                    "{\"source_sheet\":\"Stage1Test\",\"source_address\":\"O1:T13\",\"dashboard_sheet_name\":\"Agent下拉看板\",\"title\":\"销售经营看板（兼容）\",\"date_field\":\"日期\",\"category_field\":\"产品\",\"series_field\":\"区域\",\"value_field\":\"销售额\",\"filter_fields\":[\"区域\",\"类别\"],\"filter_mode\":\"dropdown\",\"aggregation\":\"sum\",\"top_n\":5,\"number_format\":\"¥#,##0\"}");
                InvokeOperation(
                    "AgentForExcel.Operations.Cell.WriteRangeOp+Factory",
                    "{\"sheet\":\"Stage1Test\",\"address\":\"AA1\",\"values\":[[\"产品\",\"数量\",\"区域\"],[\" A \",\"10\",\" 华东 \"],[\" A \",\"10\",\" 华东 \"],[null,null,null],[\"B\",\"20\",\"华南\"],[\"C\",\"30\",\"华北\"]]}");
                var powerQueryCreated = InvokeOperation(
                    "AgentForExcel.Operations.PowerQuery.CreateRangeQueryOp+Factory",
                    "{\"source_sheet\":\"Stage1Test\",\"source_address\":\"AA1:AC6\",\"query_name\":\"SmokePQ\",\"remove_blank_rows\":true,\"trim_text\":true,\"remove_duplicates\":true,\"rename_columns\":[{\"from\":\"产品\",\"to\":\"产品名\"}],\"column_types\":[{\"field\":\"数量\",\"type\":\"integer\"}],\"select_columns\":[\"产品名\",\"数量\",\"区域\"]}");
                var powerQueryList = InvokeOperation(
                    "AgentForExcel.Operations.PowerQuery.ListQueriesOp+Factory", "{}");
                var powerQueryLoaded = InvokeOperation(
                    "AgentForExcel.Operations.PowerQuery.LoadQueryOp+Factory",
                    "{\"query_name\":\"SmokePQ\",\"destination_sheet\":\"PQ结果\",\"destination_address\":\"A1\"}");
                InvokeOperation(
                    "AgentForExcel.Operations.Cell.WriteRangeOp+Factory",
                    "{\"sheet\":\"Stage1Test\",\"address\":\"AE1\",\"values\":[[\"产品名\",\"类别\"],[\"A\",\"核心\"],[\"B\",\"成长\"],[\"C\",\"核心\"]]}");
                InvokeOperation(
                    "AgentForExcel.Operations.PowerQuery.CreateRangeQueryOp+Factory",
                    "{\"source_sheet\":\"Stage1Test\",\"source_address\":\"AE1:AF4\",\"query_name\":\"DimProduct\",\"remove_blank_rows\":true,\"trim_text\":true,\"remove_duplicates\":true}");

                Assert(Convert.ToString(worksheet.Range("A2").Value2) == "A", "A2 写值失败");
                Assert(Convert.ToDouble(worksheet.Range("B3").Value2) == 20, "B3 写值失败");
                Assert(Convert.ToDouble(worksheet.Range("C2").Value2) == 20, "C2 公式结果失败");
                Assert(Convert.ToDouble(worksheet.Range("C3").Value2) == 40, "C3 公式结果失败");
                Assert(powerQueryCreated.Contains("源数据未修改"), "Power Query 创建结果没有声明源表保护");
                Assert(powerQueryList.StartsWith("__AGENT_PQ_LIST__"), "Power Query 列表没有返回结构化结果");
                using (var queryListJson = JsonDocument.Parse(powerQueryList.Substring("__AGENT_PQ_LIST__".Length)))
                {
                    Assert(queryListJson.RootElement.GetProperty("count").GetInt32() >= 1, "Power Query 列表为空");
                    Assert(queryListJson.RootElement.GetProperty("queries")[0].GetProperty("name").GetString() == "SmokePQ", "Power Query 名称不正确");
                }
                Assert(powerQueryLoaded.Contains("3 行 × 3 列"), "Power Query 加载后的数据规模不正确：" + powerQueryLoaded);
                dynamic powerQuerySheet = workbook.Worksheets.Item("PQ结果");
                Assert(Convert.ToString(powerQuerySheet.Range("A2").Value2) == "A", "Power Query 文本去空格失败");
                Assert(Convert.ToDouble(powerQuerySheet.Range("B2").Value2) == 10, "Power Query 数值类型转换失败");
                Assert(Convert.ToString(powerQuerySheet.Range("C2").Value2) == "华东", "Power Query 区域去空格失败");
                Assert(Convert.ToString(powerQuerySheet.Range("A4").Value2) == "C", "Power Query 去重或空行清理失败");
                worksheet.Range("AB6").Value2 = 35;
                var powerQueryRefreshed = InvokeOperation(
                    "AgentForExcel.Operations.PowerQuery.RefreshQueryOp+Factory",
                    "{\"query_name\":\"SmokePQ\"}");
                Assert(powerQueryRefreshed.Contains("1 个加载结果"), "Power Query 刷新没有命中加载表");
                Assert(Convert.ToDouble(powerQuerySheet.Range("B4").Value2) == 35, "Power Query 刷新后结果没有同步源数据变化");
                ReleaseCom(powerQuerySheet);
                Assert(Convert.ToBoolean(worksheet.Range("A1").Font.Bold), "标题加粗失败");
                Assert(read.StartsWith("__AGENT_TABLE_PREVIEW__"), "读取结果没有返回结构化表格数据");
                using (var preview = JsonDocument.Parse(read.Substring("__AGENT_TABLE_PREVIEW__".Length)))
                {
                    Assert(preview.RootElement.GetProperty("headers").GetArrayLength() == 3, "表格列标题数量错误");
                    Assert(preview.RootElement.GetProperty("rows").GetArrayLength() == 3, "表格预览行数错误");
                }
                Assert(profile.StartsWith("__AGENT_DATA_PROFILE__"), "数据体检没有返回结构化结果");
                using (var profileJson = JsonDocument.Parse(profile.Substring("__AGENT_DATA_PROFILE__".Length)))
                {
                    Assert(profileJson.RootElement.GetProperty("data_rows").GetInt32() == 6, "数据体检行数错误");
                    Assert(profileJson.RootElement.GetProperty("fields")[0].GetProperty("InferredType").GetString() == "date", "季度字段没有识别为时间维度");
                    Assert(profileJson.RootElement.GetProperty("fields")[1].GetProperty("Role").GetString() == "measure", "Volume 没有识别为数值指标");
                }
                Assert(Convert.ToInt32(worksheet.ChartObjects().Count) == 5, "报告级图表组创建失败");
                Assert(Convert.ToBoolean(worksheet.ChartObjects(2).Chart.SeriesCollection(1).HasDataLabels), "条形图没有报告标签");
                Assert(Convert.ToInt32(worksheet.ChartObjects(3).Chart.SeriesCollection(1).MarkerSize) == 7, "折线图标记点样式错误");
                Assert(Convert.ToDouble(worksheet.ChartObjects(4).Chart.SeriesCollection(1).Format.Fill.Transparency) > 0.2, "面积图透明度样式错误");
                Assert(Convert.ToInt32(worksheet.ChartObjects(5).Chart.SeriesCollection(1).MarkerSize) == 8, "散点图标记点样式错误");
                Assert(Convert.ToString(worksheet.Range("G2").Value2) == "北区", "安全分析排序修改了源表行顺序");
                Assert(Convert.ToDouble(worksheet.Range("I2").Value2) == 100, "安全分析排序修改了源表数值");
                dynamic analysisSheet = workbook.Worksheets.Item("Agent分析");
                Assert(Convert.ToDouble(analysisSheet.Range("C2").Value2) == 120, "分析视图没有按销售额降序排列");
                Assert(!Convert.ToBoolean(analysisSheet.Range("C2").HasFormula), "分析视图没有转换为值快照");
                Assert(Convert.ToInt32(analysisSheet.ListObjects.Count) == 1, "分析视图没有创建格式化表格");
                ReleaseCom(analysisSheet);

                dynamic chartSheet = workbook.Worksheets.Item("Agent图表");
                Assert(Convert.ToInt32(chartSheet.ChartObjects().Count) == 1, "默认新图表工作表创建失败");
                Assert(Convert.ToInt32(chartSheet.ListObjects.Count) == 1, "图表分析页没有保留数据快照");
                Assert(Convert.ToString(chartSheet.Range("A5").Value2) == "住房", "占比图数据没有按数值降序排列");
                Assert(Convert.ToDouble(chartSheet.Range("B5").Value2) == 3000, "占比图选择了错误的数值字段");
                Assert(Convert.ToString(chartSheet.Range("A6").Value2) == "餐饮", "占比图没有排除储蓄或排序错误");
                Assert(Convert.ToInt32(chartSheet.ChartObjects(1).Chart.ChartType) == -4120, "没有创建环形图");
                var doughnutPointCount = Convert.ToInt32(chartSheet.ChartObjects(1).Chart.SeriesCollection(1).Points.Count);
                Assert(doughnutPointCount == 6, "占比图过滤后的分类数量错误，实际为 " + doughnutPointCount);
                Assert(Convert.ToBoolean(chartSheet.ChartObjects(1).Chart.SeriesCollection(1).HasDataLabels), "占比图没有数据标签");
                dynamic centerShape = chartSheet.ChartObjects(1).Chart.Shapes.Item(1);
                Assert(Convert.ToString(centerShape.TextFrame2.TextRange.Text).Contains("7,620"), "环形图中心摘要不正确");
                Assert(Convert.ToDouble(centerShape.Left) > 100 && Convert.ToDouble(centerShape.Top) > 100, "环形图中心摘要位置不正确");
                ReleaseCom(centerShape);
                var previewPath = Environment.GetEnvironmentVariable("AGENT_CHART_PREVIEW");
                if (!string.IsNullOrWhiteSpace(previewPath))
                {
                    excel.Visible = true;
                    chartSheet.Activate();
                    chartSheet.ChartObjects(1).Activate();
                    var exported = Convert.ToBoolean(chartSheet.ChartObjects(1).Chart.Export(previewPath, "PNG", false));
                    if (!exported)
                    {
                        chartSheet.ChartObjects(1).Chart.CopyPicture(1, 2, 1);
                        Application.DoEvents();
                        Thread.Sleep(200);
                        using (var preview = Clipboard.GetImage())
                        {
                            Assert(preview != null, "图表预览导出失败");
                            preview.Save(previewPath, ImageFormat.Png);
                        }
                    }
                    excel.Visible = false;
                }
                ReleaseCom(chartSheet);
                dynamic aggregatedChartSheet = workbook.Worksheets.Item("Agent图表2");
                Assert(Convert.ToInt32(aggregatedChartSheet.ChartObjects(1).Chart.SeriesCollection(1).Points.Count) == 3,
                    "重复季度没有在制图前自动聚合");
                Assert(Convert.ToDouble(aggregatedChartSheet.Range("B5").Value2) == 87150200,
                    "季度自动汇总结果错误");
                var compactAxis = aggregatedChartSheet.ChartObjects(1).Chart.Axes(2);
                Assert(Convert.ToInt32(compactAxis.DisplayUnit) == -6, "大数值纵轴没有使用百万单位");
                Assert(Convert.ToString(compactAxis.DisplayUnitLabel.Caption) == "M", "大数值纵轴单位标签错误");
                var smartChartPreview = Environment.GetEnvironmentVariable("AGENT_SMART_CHART_PREVIEW");
                if (!string.IsNullOrWhiteSpace(smartChartPreview))
                {
                    excel.Visible = true;
                    aggregatedChartSheet.Activate();
                    aggregatedChartSheet.ChartObjects(1).Activate();
                    var smartChartExported = Convert.ToBoolean(aggregatedChartSheet.ChartObjects(1).Chart.Export(smartChartPreview, "PNG", false));
                    if (!smartChartExported)
                    {
                        aggregatedChartSheet.ChartObjects(1).Chart.CopyPicture(1, 2, 1);
                        Application.DoEvents();
                        Thread.Sleep(200);
                        using (var preview = Clipboard.GetImage())
                        {
                            Assert(preview != null, "智能图表预览导出失败");
                            preview.Save(smartChartPreview, ImageFormat.Png);
                        }
                    }
                    excel.Visible = false;
                }
                ReleaseCom(aggregatedChartSheet);
                dynamic pivotSheet = workbook.Worksheets.Item("PivotTest");
                Assert(Convert.ToInt32(pivotSheet.PivotTables().Count) == 1, "数据透视表创建失败");
                ReleaseCom(pivotSheet);

                Assert(Convert.ToString(worksheet.Range("O2").Value2) == dashboardSourceDateBefore, "联动看板修改了源数据");
                Assert(Convert.ToDouble(worksheet.Range("S2").Value2) == dashboardSourceValueBefore, "联动看板修改了源数值");
                dynamic dashboardSheet = workbook.Worksheets.Item("Agent看板");
                Assert(Convert.ToInt32(dashboardSheet.ChartObjects().Count) == 3, "联动看板图表数量错误");
                for (var chartIndex = 1; chartIndex <= 3; chartIndex++)
                {
                    dynamic pivotLayout = dashboardSheet.ChartObjects(chartIndex).Chart.PivotLayout;
                    Assert(pivotLayout != null, "看板图表 " + chartIndex + " 不是动态透视图");
                    ReleaseCom(pivotLayout);
                }
                Assert(Convert.ToInt32(dashboardSheet.PivotTables().Count) == 1, "联动看板明细透视表缺失");
                Assert(Convert.ToBoolean(dashboardSheet.Range("A11").HasFormula), "联动看板 KPI 没有连接透视指标");
                var dashboardKpiRaw = dashboardSheet.Range("A11").Value2;
                double dashboardKpiValue;
                Assert(double.TryParse(Convert.ToString(dashboardKpiRaw), NumberStyles.Any, CultureInfo.CurrentCulture, out dashboardKpiValue) && dashboardKpiValue > 0,
                    "联动看板 KPI 计算失败，公式=" + Convert.ToString(dashboardSheet.Range("A11").Formula) +
                    "，值=" + Convert.ToString(dashboardKpiRaw));
                dynamic dashboardDataSheet = workbook.Worksheets.Item("Agent看板数据");
                Assert(Convert.ToInt32(dashboardDataSheet.Visible) == 2, "联动看板辅助数据页没有设为深度隐藏");
                Assert(Convert.ToInt32(dashboardDataSheet.PivotTables().Count) == 6, "联动看板共享透视数据模型不完整");
                Assert(Convert.ToInt32(workbook.SlicerCaches.Count) == 2, "联动看板全局筛选器数量错误");
                for (var cacheIndex = 1; cacheIndex <= Convert.ToInt32(workbook.SlicerCaches.Count); cacheIndex++)
                {
                    var linkedPivotCount = Convert.ToInt32(workbook.SlicerCaches.Item(cacheIndex).PivotTables.Count);
                    Assert(linkedPivotCount == 7,
                        "全局筛选器没有联动全部 KPI、图表和明细透视表，切片器 " + cacheIndex + " 实际连接 " + linkedPivotCount + " 张");
                }
                var nativeRankingPointsBefore = Convert.ToInt32(dashboardSheet.ChartObjects(1).Chart.SeriesCollection(1).Points.Count);
                Assert(nativeRankingPointsBefore == 4, "原生看板初始产品数量错误");
                dynamic regionSlicerCache = SelectSingleSlicerItem(workbook, "区域", "华东");
                Application.DoEvents();
                Thread.Sleep(250);
                Assert(Math.Abs(Convert.ToDouble(dashboardSheet.Range("A11").Value2) - 59100d) < 0.001,
                    "原生切片器选择华东后 KPI 没有联动，实际=" + Convert.ToString(dashboardSheet.Range("A11").Value2));
                Assert(Convert.ToInt32(dashboardSheet.ChartObjects(1).Chart.SeriesCollection(1).Points.Count) == 3,
                    "原生切片器选择华东后排名图没有动态缩减为 3 个产品");
                regionSlicerCache.ClearManualFilter();
                Application.DoEvents();
                ReleaseCom(regionSlicerCache);
                ReleaseCom(dashboardDataSheet);
                ReleaseCom(dashboardSheet);

                dynamic dropdownSheet = workbook.Worksheets.Item("Agent下拉看板");
                dynamic dropdownDataSheet = workbook.Worksheets.Item("Agent下拉看板数据");
                Assert(Convert.ToInt32(dropdownSheet.ChartObjects().Count) == 3, "下拉兼容看板图表数量错误");
                Assert(Convert.ToInt32(workbook.SlicerCaches.Count) == 2, "下拉兼容模式不应创建额外切片器");
                Assert(Convert.ToInt32(dropdownSheet.Range("A7").Validation.Type) == 3, "区域下拉列表没有创建");
                Assert(Convert.ToInt32(dropdownSheet.Range("E7").Validation.Type) == 3, "类别下拉列表没有创建");
                Assert(Convert.ToString(dropdownSheet.Range("A7").Value2) == "（全部）", "下拉兼容模式默认值错误");
                dropdownSheet.Range("A7").Value2 = "华东";
                Application.DoEvents();
                Thread.Sleep(250);
                Assert(Math.Abs(Convert.ToDouble(dropdownSheet.Range("A11").Value2) - 59100d) < 0.001,
                    "区域下拉选择华东后 KPI 没有联动，实际=" + Convert.ToString(dropdownSheet.Range("A11").Value2));
                Assert(Convert.ToInt32(dropdownSheet.ChartObjects(1).Chart.SeriesCollection(1).Points.Count) == 3,
                    "区域下拉选择华东后排名图没有动态缩减为 3 个产品");
                var interactionManagerType = _agentAssembly.GetType(
                    "AgentForExcel.Operations.Dashboard.DashboardInteractionManager", true);
                interactionManagerType.GetMethod("Shutdown", BindingFlags.Static | BindingFlags.NonPublic)
                    .Invoke(null, null);
                interactionManagerType.GetMethod("Initialize", BindingFlags.Static | BindingFlags.NonPublic)
                    .Invoke(null, new object[] { excel });
                dropdownSheet.Range("E7").Value2 = "办公";
                Application.DoEvents();
                Thread.Sleep(250);
                Assert(Math.Abs(Convert.ToDouble(dropdownSheet.Range("A11").Value2) - 12800d) < 0.001,
                    "区域＋类别组合下拉没有联动，实际=" + Convert.ToString(dropdownSheet.Range("A11").Value2));
                dynamic dropdownPivotLayout = dropdownSheet.ChartObjects(1).Chart.PivotLayout;
                Assert(dropdownPivotLayout != null, "下拉兼容看板没有继续使用动态透视图");
                ReleaseCom(dropdownPivotLayout);
                Assert(Convert.ToInt32(dropdownDataSheet.Visible) == 2, "下拉兼容看板辅助页没有深度隐藏");
                Assert(!Convert.ToBoolean(workbook.HasVBProject), "看板不应向工作簿注入 VBA 工程");
                ReleaseCom(dropdownDataSheet);
                ReleaseCom(dropdownSheet);

                var dangerousFormulaBlocked = false;
                try
                {
                    InvokeOperation(
                        "AgentForExcel.Operations.Cell.FillFormulaOp+Factory",
                        "{\"sheet\":\"Stage1Test\",\"address\":\"D1\",\"formula\":\"=WEBSERVICE(\\\"https://example.com\\\")\"}");
                }
                catch (TargetInvocationException ex)
                {
                    dangerousFormulaBlocked = (ex.InnerException?.Message ?? ex.Message).Contains("WEBSERVICE");
                }
                Assert(dangerousFormulaBlocked, "危险公式没有被拦截");

                var oversizedRangeBlocked = false;
                try
                {
                    InvokeOperation(
                        "AgentForExcel.Operations.Cell.FormatRangeOp+Factory",
                        "{\"sheet\":\"Stage1Test\",\"address\":\"A1:A50001\",\"bold\":true}");
                }
                catch (TargetInvocationException ex)
                {
                    oversizedRangeBlocked = (ex.InnerException?.Message ?? ex.Message).Contains("最多允许");
                }
                Assert(oversizedRangeBlocked, "超大区域没有被拦截");

                TestDispatcherConfirmation(worksheet);
                TestAgentLoop();
                TestEmptyReplyRecovery();
                TestPrematurePlanReplyRecovery();
                TestTaskPlanCompletionGuard();
                TestWriteVerificationGuard();
                TestStreamingAccumulator();
                TestSettingsAndConversationPersistence();
                ExportUiPreviewsIfRequested();
                var dashboardWorkbookPath = Environment.GetEnvironmentVariable("AGENT_DASHBOARD_WORKBOOK");
                if (!string.IsNullOrWhiteSpace(dashboardWorkbookPath))
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(dashboardWorkbookPath));
                    workbook.SaveAs(dashboardWorkbookPath, 51);
                }

                Console.WriteLine("PASS");
                Console.WriteLine(write);
                Console.WriteLine(formula);
                Console.WriteLine(format);
                Console.WriteLine(chart);
                Console.WriteLine(barChart);
                Console.WriteLine(lineChart);
                Console.WriteLine(areaChart);
                Console.WriteLine(scatterChart);
                Console.WriteLine(analysisView);
                Console.WriteLine(doughnut);
                Console.WriteLine(profile);
                Console.WriteLine(aggregatedTrend);
                Console.WriteLine(pivot);
                Console.WriteLine(dashboard);
                Console.WriteLine(dropdownDashboard);
                Console.WriteLine("危险公式拦截：通过");
                Console.WriteLine("区域上限拦截：通过");
                Console.WriteLine("写操作确认回调：通过");
                Console.WriteLine("多轮 Agent 工具回传：通过");
                Console.WriteLine("空回复自动续写：通过");
                Console.WriteLine("计划式回复防提前结束：通过");
                RunPowerPivotSmoke();
                Console.WriteLine("任务计划状态与完成守卫：通过");
                Console.WriteLine("写入结果强制回读验收：通过");
                Console.WriteLine("Power Query 创建、加载与刷新：通过");
                Console.WriteLine("Power Pivot 关系、DAX 与模型透视：通过");
                Console.WriteLine("流式文本与工具调用拼接：通过");
                Console.WriteLine("结构化表格预览：通过");
                Console.WriteLine("源表只读与安全分析视图：通过");
                Console.WriteLine("现代环形图与正确字段选择：通过");
                Console.WriteLine("多模型与多会话持久化：通过");
                Console.WriteLine("条形/折线/面积/散点报告级样式：通过");
                Console.WriteLine("原生切片器真实筛选与动态透视图：通过");
                Console.WriteLine("下拉兼容筛选、组合条件与重启恢复：通过");
                Console.WriteLine("数据体检、重复横轴聚合与紧凑数值轴：通过");
                return 0;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("FAIL");
                Console.Error.WriteLine(ex);
                return 1;
            }
            finally
            {
                try { if (workbook != null) workbook.Close(false); } catch { }
                try { if (excel != null) excel.Quit(); } catch { }
                ReleaseCom(worksheet);
                ReleaseCom(workbook);
                ReleaseCom(excel);
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private static int RunSettingsAndCatalogSmoke()
        {
            var path = Path.Combine(Path.GetTempPath(), "agent-settings-" + Guid.NewGuid().ToString("N") + ".json");
            try
            {
                var settingsType = _agentAssembly.GetType("AgentForExcel.Models.UserSettings", true);
                var settings = Activator.CreateInstance(settingsType);
                settingsType.GetProperty("PreserveSourceData").SetValue(settings, true);
                settingsType.GetProperty("PreferNewWorksheetForOutputs").SetValue(settings, true);
                settingsType.GetProperty("AutomationMode").SetValue(settings, "custom");
                settingsType.GetProperty("AutoAllowNewSheetOutputs").SetValue(settings, true);
                settingsType.GetProperty("AutoAllowSelectedBlankWrites").SetValue(settings, false);
                settingsType.GetProperty("AutoWriteMaxCells").SetValue(settings, 1200);
                settingsType.GetProperty("EnablePowerQuery").SetValue(settings, false);
                settingsType.GetProperty("DefaultAnalysisScope").SetValue(settings, "Selection");
                settingsType.GetMethod("SaveTo").Invoke(settings, new object[] { path });

                var loaded = settingsType.GetMethod("LoadFrom", BindingFlags.Public | BindingFlags.Static)
                    .Invoke(null, new object[] { path });
                Assert(Convert.ToBoolean(settingsType.GetProperty("PreserveSourceData").GetValue(loaded)), "源数据保护设置未持久化");
                Assert(Convert.ToBoolean(settingsType.GetProperty("PreferNewWorksheetForOutputs").GetValue(loaded)), "新工作表输出设置未持久化");
                Assert(Convert.ToString(settingsType.GetProperty("AutomationMode").GetValue(loaded)) == "custom", "自动化模式未持久化");
                Assert(!Convert.ToBoolean(settingsType.GetProperty("AutoAllowSelectedBlankWrites").GetValue(loaded)), "自动写入白名单未持久化");
                Assert(Convert.ToInt32(settingsType.GetProperty("AutoWriteMaxCells").GetValue(loaded)) == 1200, "自动写入上限未持久化");
                Assert(!Convert.ToBoolean(settingsType.GetProperty("EnablePowerQuery").GetValue(loaded)), "工具权限设置未持久化");
                Assert(Convert.ToString(settingsType.GetProperty("DefaultAnalysisScope").GetValue(loaded)) == "Selection", "默认范围设置未持久化");

                var promptType = _agentAssembly.GetType("AgentForExcel.AI.PromptBuilder", true);
                var prompt = Convert.ToString(promptType.GetMethod("BuildSystemPrompt").Invoke(null, new[] { loaded }));
                Assert(prompt.Contains("当前选区"), "运行时提示词未应用默认范围");
                Assert(prompt.Contains("自定义白名单"), "运行时提示词未应用自动化权限");
                Assert(prompt.Contains("禁用，不得调用 pq_ 工具"), "运行时提示词未应用工具权限");

                var catalogType = _agentAssembly.GetType("AgentForExcel.Services.CapabilityCatalog", true);
                var items = (IEnumerable)catalogType.GetProperty("Items", BindingFlags.Public | BindingFlags.Static).GetValue(null);
                var count = 0;
                var categories = new HashSet<string>();
                foreach (var item in items)
                {
                    count++;
                    categories.Add(Convert.ToString(item.GetType().GetProperty("Category").GetValue(item)));
                }
                Assert(count >= 15, "功能中心能力数量不足");
                Assert(categories.Contains("分析数据") && categories.Contains("生成报告") &&
                       categories.Contains("数据工程") && categories.Contains("自动化"), "功能中心分类不完整");

                if (System.Windows.Application.Current == null)
                    new System.Windows.Application();
                var settingsWindowType = _agentAssembly.GetType("AgentForExcel.UI.SettingsWindow", true);
                var settingsWindow = (System.Windows.Window)Activator.CreateInstance(
                    settingsWindowType, new object[] { loaded, "workbook" });
                Assert(settingsWindow.Title.Contains("设置"), "设置窗口未正确加载");
                settingsWindow.Close();

                var chatViewType = _agentAssembly.GetType("AgentForExcel.UI.ChatView", true);
                var chatView = Activator.CreateInstance(chatViewType);
                Assert(chatView != null, "聊天工作区未正确加载");

                Console.WriteLine("PASS");
                Console.WriteLine("设置持久化、运行时提示词、功能中心目录和 WPF 界面加载验证通过。能力数=" + count);
                return 0;
            }
            catch (Exception ex)
            {
                Console.WriteLine("FAIL");
                Console.WriteLine(ex);
                return 1;
            }
            finally
            {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        private static void RunPowerPivotSmoke()
        {
            dynamic modelExcel = null;
            dynamic modelWorkbook = null;
            dynamic sheet = null;
            var originalContext = _context;
            try
            {
                modelExcel = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application", true));
                modelExcel.AutomationSecurity = 1;
                modelExcel.Visible = false;
                modelExcel.DisplayAlerts = false;
                modelWorkbook = modelExcel.Workbooks.Add();
                modelWorkbook.Activate();
                Assert(Convert.ToString(modelExcel.ActiveWorkbook.Name) == Convert.ToString(modelWorkbook.Name),
                    "Power Pivot 测试工作簿没有成为活动工作簿");
                _context = CreateContext(modelExcel);
                sheet = modelWorkbook.Worksheets.Item(1);
                sheet.Name = "ModelTest";
                sheet.Range("A1:C4").Value2 = new object[,]
                {
                    { "产品名", "数量", "区域" },
                    { "A", 10, "华东" },
                    { "B", 20, "华南" },
                    { "C", 35, "华北" }
                };
                sheet.Range("E1:F4").Value2 = new object[,]
                {
                    { "产品名", "类别" },
                    { "A", "核心" },
                    { "B", "成长" },
                    { "C", "核心" }
                };

                modelWorkbook.Queries.Add("SmokePQ",
                    "let Source = #table({\"产品名\",\"数量\",\"区域\"}, {{\"A\",10,\"华东\"},{\"B\",20,\"华南\"},{\"C\",35,\"华北\"}}) in Source",
                    "Power Pivot smoke fact query");
                modelWorkbook.Queries.Add("DimProduct",
                    "let Source = #table({\"产品名\",\"类别\"}, {{\"A\",\"核心\"},{\"B\",\"成长\"},{\"C\",\"核心\"}}) in Source",
                    "Power Pivot smoke dimension query");
                var modelFact = InvokeOperation(
                    "AgentForExcel.Operations.PowerPivot.AddQueryToModelOp+Factory", "{\"query_name\":\"SmokePQ\"}");
                var modelDimension = InvokeOperation(
                    "AgentForExcel.Operations.PowerPivot.AddQueryToModelOp+Factory", "{\"query_name\":\"DimProduct\"}");
                var modelRelationship = InvokeOperation(
                    "AgentForExcel.Operations.PowerPivot.AddRelationshipOp+Factory",
                    "{\"from_table\":\"SmokePQ\",\"from_column\":\"产品名\",\"to_table\":\"DimProduct\",\"to_column\":\"产品名\"}");
                var daxMeasure = InvokeOperation(
                    "AgentForExcel.Operations.PowerPivot.AddMeasureOp+Factory",
                    "{\"table\":\"SmokePQ\",\"measure_name\":\"总数量\",\"formula\":\"SUM(SmokePQ[数量])\",\"format\":\"whole_number\",\"description\":\"清洗后数量合计\"}");
                var modelPivot = InvokeOperation(
                    "AgentForExcel.Operations.PowerPivot.CreateModelPivotOp+Factory",
                    "{\"destination_sheet\":\"模型透视\",\"destination_address\":\"A1\",\"name\":\"SmokeModelPivot\",\"rows\":[{\"table\":\"DimProduct\",\"field\":\"类别\"}],\"measures\":[\"总数量\"]}");
                var modelList = InvokeOperation(
                    "AgentForExcel.Operations.PowerPivot.ListModelOp+Factory", "{}");

                Assert(modelFact.Contains("加入数据模型") && modelDimension.Contains("加入数据模型"), "Power Pivot 模型表加载失败");
                Assert(modelRelationship.Contains("一对多关系"), "Power Pivot 关系创建失败");
                Assert(daxMeasure.Contains("总数量") && daxMeasure.Contains("SUM(SmokePQ[数量])"), "DAX 度量值创建失败");
                Assert(modelPivot.Contains("1 个 DAX 度量值"), "模型透视表创建失败");
                Assert(modelList.StartsWith("__AGENT_MODEL_LIST__"), "数据模型列表没有返回结构化结果");
                using (var modelJson = JsonDocument.Parse(modelList.Substring("__AGENT_MODEL_LIST__".Length)))
                {
                    Assert(modelJson.RootElement.GetProperty("tables").GetArrayLength() == 2, "数据模型表数量错误");
                    Assert(modelJson.RootElement.GetProperty("relationships").GetArrayLength() == 1, "数据模型关系数量错误");
                    Assert(modelJson.RootElement.GetProperty("measures").GetArrayLength() == 1, "DAX 度量值数量错误");
                }
                dynamic modelPivotSheet = modelWorkbook.Worksheets.Item("模型透视");
                Assert(Convert.ToInt32(modelPivotSheet.PivotTables().Count) == 1, "模型透视表没有创建");
                ReleaseCom(modelPivotSheet);
            }
            finally
            {
                try { if (modelWorkbook != null) modelWorkbook.Close(false); } catch { }
                try { if (modelExcel != null) modelExcel.Quit(); } catch { }
                _context = originalContext;
                ReleaseCom(sheet);
                ReleaseCom(modelWorkbook);
                ReleaseCom(modelExcel);
            }
        }

        private static object CreateContext(object excel)
        {
            var contextType = _agentAssembly.GetType("AgentForExcel.AppContext", true);
            var context = FormatterServices.GetUninitializedObject(contextType);
            var field = contextType.GetField("<Excel>k__BackingField", BindingFlags.Instance | BindingFlags.NonPublic);
            if (field == null) throw new MissingFieldException("AppContext.Excel backing field");
            field.SetValue(context, excel);
            return context;
        }

        private static string InvokeOperation(string factoryTypeName, string argumentsJson)
        {
            var factoryType = _agentAssembly.GetType(factoryTypeName, true);
            var factory = Activator.CreateInstance(factoryType);
            var operation = factoryType.GetMethod("Parse").Invoke(factory, new object[] { argumentsJson });
            var result = operation.GetType().GetMethod("Execute").Invoke(operation, new[] { _context });
            return Convert.ToString(result);
        }

        private static void TestDispatcherConfirmation(dynamic worksheet)
        {
            var contextType = _agentAssembly.GetType("AgentForExcel.AppContext", true);
            var settingsType = _agentAssembly.GetType("AgentForExcel.Models.UserSettings", true);
            var settings = Activator.CreateInstance(settingsType);
            settingsType.GetProperty("RequireConfirmOnWrite").SetValue(settings, true, null);
            settingsType.GetProperty("AutomationMode").SetValue(settings, "ask_every_time", null);
            contextType.GetField("<Settings>k__BackingField", BindingFlags.Instance | BindingFlags.NonPublic)
                .SetValue(_context, settings);

            var addInType = _agentAssembly.GetType("AgentForExcel.ThisAddIn", true);
            addInType.GetField("<App>k__BackingField", BindingFlags.Static | BindingFlags.NonPublic)
                .SetValue(null, _context);

            var dispatcherType = _agentAssembly.GetType("AgentForExcel.Operations.OperationDispatcher", true);
            var dispatcher = Activator.CreateInstance(dispatcherType);
            var factoryType = _agentAssembly.GetType("AgentForExcel.Operations.Cell.WriteRangeOp+Factory", true);
            var factory = Activator.CreateInstance(factoryType);
            dispatcherType.GetMethod("Register").Invoke(dispatcher, new[] { factory });

            var callType = _agentAssembly.GetType("AgentForExcel.Operations.OperationCall", true);
            var call = Activator.CreateInstance(callType);
            callType.GetProperty("ToolName").SetValue(call, "cell_write_range", null);
            callType.GetProperty("ArgumentsJson").SetValue(
                call,
                "{\"sheet\":\"Stage1Test\",\"address\":\"E1\",\"values\":[[\"Confirmed\"]]}",
                null);

            var listType = typeof(List<>).MakeGenericType(callType);
            var calls = (IList)Activator.CreateInstance(listType);
            calls.Add(call);

            var execute = dispatcherType.GetMethod("ExecuteAsync");
            var confirmationCount = 0;
            Func<string, bool> reject = description => { confirmationCount++; return false; };
            ((Task)execute.Invoke(dispatcher, new object[] { calls, reject })).Wait();
            Assert(confirmationCount == 1, "拒绝执行时没有触发确认回调");
            Assert(worksheet.Range("E1").Value2 == null, "拒绝后仍然写入了单元格");

            Func<string, bool> approve = description => { confirmationCount++; return true; };
            ((Task)execute.Invoke(dispatcher, new object[] { calls, approve })).Wait();
            Assert(confirmationCount == 2, "批准执行时没有触发确认回调");
            Assert(Convert.ToString(worksheet.Range("E1").Value2) == "Confirmed", "批准后没有写入单元格");
        }

        private static void TestSafeAutomationPermissions(dynamic worksheet)
        {
            var contextType = _agentAssembly.GetType("AgentForExcel.AppContext", true);
            var settingsType = _agentAssembly.GetType("AgentForExcel.Models.UserSettings", true);
            var settings = Activator.CreateInstance(settingsType);
            settingsType.GetProperty("AutomationMode").SetValue(settings, "safe_auto", null);
            settingsType.GetProperty("AutoAllowSelectedBlankWrites").SetValue(settings, true, null);
            settingsType.GetProperty("AutoWriteMaxCells").SetValue(settings, 100, null);
            contextType.GetField("<Settings>k__BackingField", BindingFlags.Instance | BindingFlags.NonPublic)
                .SetValue(_context, settings);

            worksheet.Activate();
            worksheet.Range("B2:B4").Select();
            var selectionType = _agentAssembly.GetType("AgentForExcel.Services.SelectionContextService", true);
            var selection = Activator.CreateInstance(selectionType, new object[] { worksheet.Application });
            selectionType.GetMethod("Refresh").Invoke(selection, null);
            Assert(Convert.ToBoolean(selectionType.GetMethod("LockCurrent").Invoke(selection, new object[] { "task" })),
                "无法锁定当前选区");
            var locked = selectionType.GetProperty("Locked").GetValue(selection, null);
            Assert(Convert.ToString(locked.GetType().GetProperty("Address").GetValue(locked, null)).Contains("B$2:$B$4"),
                "锁定的选区地址不正确");
            contextType.GetField("<Selection>k__BackingField", BindingFlags.Instance | BindingFlags.NonPublic)
                .SetValue(_context, selection);

            var policyType = _agentAssembly.GetType("AgentForExcel.Services.PermissionPolicyService", true);
            var policy = Activator.CreateInstance(policyType, new[] { _context });
            contextType.GetField("<Permissions>k__BackingField", BindingFlags.Instance | BindingFlags.NonPublic)
                .SetValue(_context, policy);

            var addInType = _agentAssembly.GetType("AgentForExcel.ThisAddIn", true);
            addInType.GetField("<App>k__BackingField", BindingFlags.Static | BindingFlags.NonPublic)
                .SetValue(null, _context);

            var dispatcherType = _agentAssembly.GetType("AgentForExcel.Operations.OperationDispatcher", true);
            var dispatcher = Activator.CreateInstance(dispatcherType);
            var factoryType = _agentAssembly.GetType("AgentForExcel.Operations.Cell.WriteRangeOp+Factory", true);
            dispatcherType.GetMethod("Register").Invoke(dispatcher, new[] { Activator.CreateInstance(factoryType) });
            var callType = _agentAssembly.GetType("AgentForExcel.Operations.OperationCall", true);
            var call = Activator.CreateInstance(callType);
            callType.GetProperty("ToolName").SetValue(call, "cell_write_range", null);
            callType.GetProperty("ArgumentsJson").SetValue(call,
                "{\"sheet\":\"Stage1Test\",\"address\":\"B2\",\"values\":[[1],[2],[3]]}", null);
            var listType = typeof(List<>).MakeGenericType(callType);
            var calls = (IList)Activator.CreateInstance(listType);
            calls.Add(call);

            var confirmationCount = 0;
            Func<string, bool> reject = description => { confirmationCount++; return false; };
            ((Task)dispatcherType.GetMethod("ExecuteAsync").Invoke(dispatcher, new object[] { calls, reject })).Wait();
            Assert(confirmationCount == 0, "锁定选区内空白写入不应弹出确认");
            Assert(Convert.ToDouble(worksheet.Range("B4").Value2) == 3, "安全自动写入没有执行");

            callType.GetProperty("ArgumentsJson").SetValue(call,
                "{\"sheet\":\"Stage1Test\",\"address\":\"B2\",\"values\":[[99]]}", null);
            ((Task)dispatcherType.GetMethod("ExecuteAsync").Invoke(dispatcher, new object[] { calls, reject })).Wait();
            Assert(confirmationCount == 1, "覆盖已有内容时没有触发确认");
            Assert(Convert.ToDouble(worksheet.Range("B2").Value2) == 1, "拒绝覆盖后单元格仍被修改");
        }

        private static void TestAgentLoop()
        {
            var history = new List<AgentForExcel.AI.ChatTurn>();
            var requestCount = 0;
            var executeCount = 0;

            Func<IReadOnlyList<AgentForExcel.AI.ChatTurn>, Task<AgentForExcel.AI.LlmReply>> request = turns =>
            {
                requestCount++;
                if (requestCount == 1)
                {
                    Assert(turns.Count == 1 && turns[0].Role == "user", "首轮没有写入用户消息");
                    return Task.FromResult(new AgentForExcel.AI.LlmReply
                    {
                        Text = "我先读取数据。",
                        Operations = new List<AgentForExcel.Operations.OperationCall>
                        {
                            new AgentForExcel.Operations.OperationCall
                            {
                                CallId = "call_read_1",
                                ToolName = "cell_read_range",
                                ArgumentsJson = "{\"address\":\"A1:H14\"}"
                            }
                        }
                    });
                }

                var last = turns[turns.Count - 1];
                Assert(last.Role == "tool", "工具结果没有以 role=tool 回传");
                Assert(last.ToolCallId == "call_read_1", "工具结果没有关联原 tool_call_id");
                Assert(last.Content.Contains("1 | 2 | 3"), "工具结果内容丢失");
                return Task.FromResult(new AgentForExcel.AI.LlmReply
                {
                    Text = "数据整体平稳，各行均为 1 到 8，未发现异常波动。",
                    Operations = new List<AgentForExcel.Operations.OperationCall>()
                });
            };

            Func<IReadOnlyList<AgentForExcel.Operations.OperationCall>, Task<IReadOnlyList<string>>> execute = calls =>
            {
                executeCount++;
                Assert(calls.Count == 1, "工具调用数量异常");
                return Task.FromResult<IReadOnlyList<string>>(new[] { "读取 A1:H14：1 | 2 | 3 | 4 | 5 | 6 | 7 | 8" });
            };

            var run = AgentForExcel.AI.AgentLoopRunner.RunAsync(
                "分析当前表格趋势",
                history,
                request,
                execute,
                null,
                null,
                null).GetAwaiter().GetResult();

            Assert(run.Completed, "Agent 循环没有正常完成");
            Assert(run.Rounds == 2, "Agent 没有在工具执行后自动发起第二轮");
            Assert(requestCount == 2, "模型请求次数不正确");
            Assert(executeCount == 1, "工具被重复执行");
            Assert(history[history.Count - 1].Role == "assistant", "最终回答没有进入历史");
        }

        private static void TestStreamingAccumulator()
        {
            var accumulator = new AgentForExcel.AI.StreamingReplyAccumulator();
            var first = accumulator.Consume("{\"choices\":[{\"delta\":{\"content\":\"正在\"}}]}");
            var second = accumulator.Consume("{\"choices\":[{\"delta\":{\"content\":\"分析\",\"tool_calls\":[{\"index\":0,\"id\":\"call_1\",\"function\":{\"name\":\"cell_read_\",\"arguments\":\"{\\\"address\\\":\"}}]}}]}");
            accumulator.Consume("{\"choices\":[{\"delta\":{\"tool_calls\":[{\"index\":0,\"function\":{\"name\":\"range\",\"arguments\":\"\\\"A1:H12\\\"}\"}}]}}]}");
            var reply = accumulator.BuildReply();

            Assert(first == "正在" && second == "分析", "流式文本增量解析失败");
            Assert(reply.Text == "正在分析", "流式文本拼接失败");
            Assert(reply.Operations.Count == 1, "流式工具调用数量错误");
            Assert(reply.Operations[0].CallId == "call_1", "流式 tool_call_id 丢失");
            Assert(reply.Operations[0].ToolName == "cell_read_range", "流式工具名拼接失败");
            Assert(reply.Operations[0].ArgumentsJson == "{\"address\":\"A1:H12\"}", "流式参数拼接失败");
        }

        private static void TestEmptyReplyRecovery()
        {
            var history = new List<AgentForExcel.AI.ChatTurn>();
            var requestCount = 0;
            Func<IReadOnlyList<AgentForExcel.AI.ChatTurn>, Task<AgentForExcel.AI.LlmReply>> request = turns =>
            {
                requestCount++;
                return Task.FromResult(new AgentForExcel.AI.LlmReply
                {
                    Text = requestCount == 1 ? string.Empty : "这是补发的最终答复。",
                    Operations = new List<AgentForExcel.Operations.OperationCall>()
                });
            };
            Func<IReadOnlyList<AgentForExcel.Operations.OperationCall>, Task<IReadOnlyList<string>>> execute = calls =>
                Task.FromResult<IReadOnlyList<string>>(new string[0]);

            var run = AgentForExcel.AI.AgentLoopRunner.RunAsync(
                "请分析数据",
                history,
                request,
                execute,
                null,
                null,
                null).GetAwaiter().GetResult();

            Assert(run.Completed && run.Rounds == 2, "空回复后没有自动补发最终答复");
            Assert(requestCount == 2, "空回复恢复请求次数错误");
        }

        private static void TestPrematurePlanReplyRecovery()
        {
            var history = new List<AgentForExcel.AI.ChatTurn>();
            var requestCount = 0;
            var executeCount = 0;
            Func<IReadOnlyList<AgentForExcel.AI.ChatTurn>, Task<AgentForExcel.AI.LlmReply>> request = turns =>
            {
                requestCount++;
                if (requestCount == 1)
                {
                    return Task.FromResult(new AgentForExcel.AI.LlmReply
                    {
                        Text = "我需要先读取当前工作表的数据才能进行分析。让我先查看一下数据内容。",
                        Operations = new List<AgentForExcel.Operations.OperationCall>()
                    });
                }
                if (requestCount == 2)
                {
                    Assert(turns[turns.Count - 1].Content.Contains("完成检查未通过"),
                        "计划式回复后没有注入完成检查提示");
                    return Task.FromResult(new AgentForExcel.AI.LlmReply
                    {
                        Text = "正在读取数据。",
                        Operations = new List<AgentForExcel.Operations.OperationCall>
                        {
                            new AgentForExcel.Operations.OperationCall
                            {
                                CallId = "call_recovery_read",
                                ToolName = "cell_read_range",
                                ArgumentsJson = "{\"address\":\"A1:H14\"}"
                            }
                        }
                    });
                }
                return Task.FromResult(new AgentForExcel.AI.LlmReply
                {
                    Text = "已完成趋势分析，并给出主要变化和异常结论。",
                    Operations = new List<AgentForExcel.Operations.OperationCall>()
                });
            };
            Func<IReadOnlyList<AgentForExcel.Operations.OperationCall>, Task<IReadOnlyList<string>>> execute = calls =>
            {
                executeCount++;
                return Task.FromResult<IReadOnlyList<string>>(new[] { "已读取 A1:H14。" });
            };
            var run = AgentForExcel.AI.AgentLoopRunner.RunAsync(
                "分析当前工作表趋势", history, request, execute, null, null, null).GetAwaiter().GetResult();
            Assert(run.Completed && run.Rounds == 3, "计划式回复没有自动续跑到真实完成");
            Assert(run.CompletionCheckCount == 1, "完成检查次数错误");
            Assert(executeCount == 1, "恢复后工具执行次数错误");
        }

        private static void TestTaskPlanCompletionGuard()
        {
            var history = new List<AgentForExcel.AI.ChatTurn>();
            var requestCount = 0;
            Func<IReadOnlyList<AgentForExcel.AI.ChatTurn>, Task<AgentForExcel.AI.LlmReply>> request = turns =>
            {
                requestCount++;
                if (requestCount == 1)
                {
                    return Task.FromResult(new AgentForExcel.AI.LlmReply
                    {
                        Operations = new List<AgentForExcel.Operations.OperationCall>
                        {
                            new AgentForExcel.Operations.OperationCall
                            {
                                CallId = "call_plan",
                                ToolName = "task_plan",
                                ArgumentsJson = "{\"title\":\"清洗并分析\",\"steps\":[\"读取并体检数据\",\"生成分析结果\"],\"success_criteria\":[\"两步均完成并核验\"]}"
                            }
                        }
                    });
                }
                if (requestCount == 2)
                {
                    return Task.FromResult(new AgentForExcel.AI.LlmReply
                    {
                        Text = "任务已经处理好了。",
                        Operations = new List<AgentForExcel.Operations.OperationCall>()
                    });
                }
                if (requestCount == 3)
                {
                    Assert(turns[turns.Count - 1].Content.Contains("任务计划仍有未完成步骤"),
                        "未完成计划没有触发完成守卫");
                    return Task.FromResult(new AgentForExcel.AI.LlmReply
                    {
                        Operations = new List<AgentForExcel.Operations.OperationCall>
                        {
                            new AgentForExcel.Operations.OperationCall
                            {
                                CallId = "call_step_1",
                                ToolName = "task_step_update",
                                ArgumentsJson = "{\"step_index\":1,\"status\":\"completed\",\"detail\":\"字段体检通过\"}"
                            },
                            new AgentForExcel.Operations.OperationCall
                            {
                                CallId = "call_step_2",
                                ToolName = "task_step_update",
                                ArgumentsJson = "{\"step_index\":2,\"status\":\"completed\",\"detail\":\"结果已回读核验\"}"
                            }
                        }
                    });
                }
                return Task.FromResult(new AgentForExcel.AI.LlmReply
                {
                    Text = "清洗与分析已经完成，结果已核验。",
                    Operations = new List<AgentForExcel.Operations.OperationCall>()
                });
            };
            Func<IReadOnlyList<AgentForExcel.Operations.OperationCall>, Task<IReadOnlyList<string>>> execute = calls =>
            {
                var results = new List<string>();
                foreach (var call in calls)
                {
                    using (var document = JsonDocument.Parse(call.ArgumentsJson))
                    {
                        if (call.ToolName == "task_plan")
                        {
                            var steps = new List<string>();
                            foreach (var item in document.RootElement.GetProperty("steps").EnumerateArray())
                                steps.Add(item.GetString());
                            var criteria = new List<string>();
                            foreach (var item in document.RootElement.GetProperty("success_criteria").EnumerateArray())
                                criteria.Add(item.GetString());
                            results.Add(AgentForExcel.Operations.Tasking.TaskExecutionRegistry.Serialize(
                                AgentForExcel.Operations.Tasking.TaskExecutionRegistry.SetPlan(
                                    document.RootElement.GetProperty("title").GetString(), steps, criteria)));
                        }
                        else
                        {
                            results.Add(AgentForExcel.Operations.Tasking.TaskExecutionRegistry.Serialize(
                                AgentForExcel.Operations.Tasking.TaskExecutionRegistry.UpdateStep(
                                    document.RootElement.GetProperty("step_index").GetInt32(),
                                    document.RootElement.GetProperty("status").GetString(),
                                    document.RootElement.GetProperty("detail").GetString())));
                        }
                    }
                }
                return Task.FromResult<IReadOnlyList<string>>(results);
            };
            var run = AgentForExcel.AI.AgentLoopRunner.RunAsync(
                "清洗并分析当前数据", history, request, execute, null, null, null).GetAwaiter().GetResult();
            Assert(run.Completed && run.Rounds == 4, "任务计划没有持续执行到全部步骤完成");
            Assert(run.CompletionCheckCount == 1, "未完成计划的完成检查次数错误");
            Assert(AgentForExcel.Operations.Tasking.TaskExecutionRegistry.IsComplete,
                "任务计划最终状态不是完成");
        }

        private static void TestWriteVerificationGuard()
        {
            var history = new List<AgentForExcel.AI.ChatTurn>();
            var requestCount = 0;
            Func<IReadOnlyList<AgentForExcel.AI.ChatTurn>, Task<AgentForExcel.AI.LlmReply>> request = turns =>
            {
                requestCount++;
                if (requestCount == 1)
                {
                    return Task.FromResult(new AgentForExcel.AI.LlmReply
                    {
                        Operations = new List<AgentForExcel.Operations.OperationCall>
                        {
                            new AgentForExcel.Operations.OperationCall
                            {
                                CallId = "call_write_verify",
                                ToolName = "cell_write_range",
                                ArgumentsJson = "{}"
                            }
                        }
                    });
                }
                if (requestCount == 2)
                {
                    return Task.FromResult(new AgentForExcel.AI.LlmReply
                    {
                        Text = "已经写入完成。",
                        Operations = new List<AgentForExcel.Operations.OperationCall>()
                    });
                }
                if (requestCount == 3)
                {
                    Assert(turns[turns.Count - 1].Content.Contains("回读核验"), "写入后没有触发回读验收提示");
                    return Task.FromResult(new AgentForExcel.AI.LlmReply
                    {
                        Operations = new List<AgentForExcel.Operations.OperationCall>
                        {
                            new AgentForExcel.Operations.OperationCall
                            {
                                CallId = "call_read_verify",
                                ToolName = "cell_read_range",
                                ArgumentsJson = "{}"
                            }
                        }
                    });
                }
                return Task.FromResult(new AgentForExcel.AI.LlmReply
                {
                    Text = "写入内容已经回读核验，结果正确。",
                    Operations = new List<AgentForExcel.Operations.OperationCall>()
                });
            };
            Func<IReadOnlyList<AgentForExcel.Operations.OperationCall>, Task<IReadOnlyList<string>>> execute = calls =>
                Task.FromResult<IReadOnlyList<string>>(new[]
                {
                    calls[0].ToolName == "cell_read_range" ? "已读取并确认目标区域。" : "已写入目标区域。"
                });
            var run = AgentForExcel.AI.AgentLoopRunner.RunAsync(
                "写入并核验数据", history, request, execute, null, null, null).GetAwaiter().GetResult();
            Assert(run.Completed && run.Rounds == 4, "写入操作没有在回读后完成");
            Assert(run.CompletionCheckCount == 1, "写入回读完成检查次数错误");
        }

        private static void TestSettingsAndConversationPersistence()
        {
            var directory = Path.Combine(Path.GetTempPath(), "AgentForExcel-Smoke-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            try
            {
                var settingsPath = Path.Combine(directory, "settings.json");
                var settingsType = _agentAssembly.GetType("AgentForExcel.Models.UserSettings", true);
                var profileType = _agentAssembly.GetType("AgentForExcel.Models.ModelProfile", true);
                var settings = Activator.CreateInstance(settingsType);
                var profiles = (IList)settingsType.GetProperty("Profiles").GetValue(settings, null);
                var secondProfile = Activator.CreateInstance(profileType);
                var secondId = Guid.NewGuid().ToString("N");
                profileType.GetProperty("Id").SetValue(secondProfile, secondId, null);
                profileType.GetProperty("DisplayName").SetValue(secondProfile, "GLM 分析", null);
                profileType.GetProperty("ProviderName").SetValue(secondProfile, "智谱 GLM（国内）", null);
                profileType.GetProperty("ApiKey").SetValue(secondProfile, "test-key", null);
                profileType.GetProperty("BaseUrl").SetValue(secondProfile, "https://open.bigmodel.cn/api/paas/v4", null);
                profileType.GetProperty("Model").SetValue(secondProfile, "glm-4.7", null);
                profileType.GetProperty("Temperature").SetValue(secondProfile, 0.2, null);

                var genericListType = typeof(List<>).MakeGenericType(profileType);
                var profileList = (IList)Activator.CreateInstance(genericListType);
                profileList.Add(profiles[0]);
                profileList.Add(secondProfile);
                settingsType.GetMethod("ReplaceProfiles").Invoke(settings, new object[] { profileList, secondId });
                settingsType.GetMethod("SaveTo").Invoke(settings, new object[] { settingsPath });

                var loaded = settingsType.GetMethod("LoadFrom", BindingFlags.Public | BindingFlags.Static)
                    .Invoke(null, new object[] { settingsPath });
                var loadedProfiles = (IList)settingsType.GetProperty("Profiles").GetValue(loaded, null);
                Assert(loadedProfiles.Count == 2, "多模型配置没有持久化");
                Assert(Convert.ToString(settingsType.GetProperty("ActiveProfileId").GetValue(loaded, null)) == secondId,
                    "活动模型没有持久化");
                Assert(Convert.ToString(settingsType.GetProperty("Model").GetValue(loaded, null)) == "glm-4.7",
                    "活动模型切换失败");

                var chatPath = Path.Combine(directory, "conversations.json");
                var storeType = _agentAssembly.GetType("AgentForExcel.Services.ChatHistoryStore", true);
                var store = Activator.CreateInstance(storeType, new object[] { chatPath });
                var document = storeType.GetMethod("Load").Invoke(store, null);
                var conversation = storeType.GetMethod("CreateConversation").Invoke(store, new[] { document, secondId });
                var conversationType = conversation.GetType();
                conversationType.GetProperty("Title").SetValue(conversation, "预算分析", null);
                var history = (IList)conversationType.GetProperty("History").GetValue(conversation, null);
                var turnType = _agentAssembly.GetType("AgentForExcel.AI.ChatTurn", true);
                var turn = Activator.CreateInstance(turnType);
                turnType.GetProperty("Role").SetValue(turn, "user", null);
                turnType.GetProperty("Content").SetValue(turn, "分析预算", null);
                history.Add(turn);

                var operationCallType = _agentAssembly.GetType("AgentForExcel.Operations.OperationCall", true);
                var operationListType = typeof(List<>).MakeGenericType(operationCallType);
                var operationList = (IList)Activator.CreateInstance(operationListType);
                var operationCall = Activator.CreateInstance(operationCallType);
                operationCallType.GetProperty("CallId").SetValue(operationCall, "call_saved_1", null);
                operationCallType.GetProperty("ToolName").SetValue(operationCall, "cell_read_range", null);
                operationCallType.GetProperty("ArgumentsJson").SetValue(operationCall, "{\"address\":\"A1:C9\"}", null);
                operationList.Add(operationCall);
                var assistantTurn = Activator.CreateInstance(turnType);
                turnType.GetProperty("Role").SetValue(assistantTurn, "assistant", null);
                turnType.GetProperty("Content").SetValue(assistantTurn, "我先读取数据。", null);
                turnType.GetProperty("ToolCalls").SetValue(assistantTurn, operationList, null);
                history.Add(assistantTurn);
                var toolTurn = Activator.CreateInstance(turnType);
                turnType.GetProperty("Role").SetValue(toolTurn, "tool", null);
                turnType.GetProperty("ToolCallId").SetValue(toolTurn, "call_saved_1", null);
                turnType.GetProperty("Content").SetValue(toolTurn, "已读取 A1:C9", null);
                history.Add(toolTurn);

                var messages = (IList)conversationType.GetProperty("Messages").GetValue(conversation, null);
                var persistedMessageType = _agentAssembly.GetType("AgentForExcel.Models.PersistedChatMessage", true);
                var persistedMessage = Activator.CreateInstance(persistedMessageType);
                persistedMessageType.GetProperty("Role").SetValue(persistedMessage,
                    Enum.Parse(_agentAssembly.GetType("AgentForExcel.Models.ChatRole", true), "User"), null);
                persistedMessageType.GetProperty("Kind").SetValue(persistedMessage,
                    Enum.Parse(_agentAssembly.GetType("AgentForExcel.Models.ChatMessageKind", true), "Text"), null);
                persistedMessageType.GetProperty("Text").SetValue(persistedMessage, "分析预算", null);
                messages.Add(persistedMessage);
                storeType.GetMethod("Save").Invoke(store, new[] { document });

                var reloadedDocument = storeType.GetMethod("Load").Invoke(store, null);
                var conversations = (IList)reloadedDocument.GetType().GetProperty("Conversations").GetValue(reloadedDocument, null);
                Assert(conversations.Count == 1, "对话列表没有持久化");
                var reloadedConversation = conversations[0];
                Assert(Convert.ToString(reloadedConversation.GetType().GetProperty("Title").GetValue(reloadedConversation, null)) == "预算分析",
                    "对话标题没有持久化");
                var reloadedHistory = (IList)reloadedConversation.GetType().GetProperty("History").GetValue(reloadedConversation, null);
                Assert(reloadedHistory.Count == 3, "协议级对话上下文没有持久化");
                var reloadedToolCalls = (IList)turnType.GetProperty("ToolCalls").GetValue(reloadedHistory[1], null);
                Assert(reloadedToolCalls.Count == 1 &&
                       Convert.ToString(operationCallType.GetProperty("CallId").GetValue(reloadedToolCalls[0], null)) == "call_saved_1",
                    "工具调用上下文没有持久化");
                var reloadedMessages = (IList)reloadedConversation.GetType().GetProperty("Messages").GetValue(reloadedConversation, null);
                Assert(reloadedMessages.Count == 1, "可见聊天记录没有持久化");
            }
            finally
            {
                try { Directory.Delete(directory, true); } catch { }
            }
        }

        private static dynamic SelectSingleSlicerItem(dynamic workbook, string fieldName, string itemName)
        {
            for (var cacheIndex = 1; cacheIndex <= Convert.ToInt32(workbook.SlicerCaches.Count); cacheIndex++)
            {
                dynamic cache = workbook.SlicerCaches.Item(cacheIndex);
                var sourceName = Convert.ToString(cache.SourceName);
                if (!string.Equals(sourceName, fieldName, StringComparison.CurrentCultureIgnoreCase))
                {
                    ReleaseCom(cache);
                    continue;
                }

                dynamic items = cache.SlicerItems;
                dynamic selected = null;
                for (var itemIndex = 1; itemIndex <= Convert.ToInt32(items.Count); itemIndex++)
                {
                    dynamic item = items.Item(itemIndex);
                    if (string.Equals(Convert.ToString(item.Name), itemName, StringComparison.CurrentCultureIgnoreCase) ||
                        string.Equals(Convert.ToString(item.Caption), itemName, StringComparison.CurrentCultureIgnoreCase))
                    {
                        selected = item;
                        break;
                    }
                    ReleaseCom(item);
                }
                Assert(selected != null, "切片器字段「" + fieldName + "」中找不到「" + itemName + "」");
                selected.Selected = true;
                for (var itemIndex = 1; itemIndex <= Convert.ToInt32(items.Count); itemIndex++)
                {
                    dynamic item = items.Item(itemIndex);
                    if (!string.Equals(Convert.ToString(item.Name), itemName, StringComparison.CurrentCultureIgnoreCase) &&
                        !string.Equals(Convert.ToString(item.Caption), itemName, StringComparison.CurrentCultureIgnoreCase))
                        item.Selected = false;
                    ReleaseCom(item);
                }
                ReleaseCom(items);
                return cache;
            }
            throw new InvalidOperationException("找不到切片器字段「" + fieldName + "」。");
        }

        private static void ExportUiPreviewsIfRequested()
        {
            var directory = Environment.GetEnvironmentVariable("AGENT_UI_PREVIEW_DIR");
            if (string.IsNullOrWhiteSpace(directory)) return;
            Directory.CreateDirectory(directory);

            var settingsType = _agentAssembly.GetType("AgentForExcel.Models.UserSettings", true);
            var settings = Activator.CreateInstance(settingsType);
            var settingsWindowType = _agentAssembly.GetType("AgentForExcel.UI.SettingsWindow", true);
            var settingsWindow = (System.Windows.Window)Activator.CreateInstance(
                settingsWindowType, new object[] { settings, "models" });
            RenderWindow(settingsWindow, Path.Combine(directory, "multi-model-settings.png"), 760, 640);
            var permissionWindow = (System.Windows.Window)Activator.CreateInstance(
                settingsWindowType, new object[] { settings, "safety" });
            RenderWindow(permissionWindow, Path.Combine(directory, "automation-permissions.png"), 760, 640);

            var chatViewType = _agentAssembly.GetType("AgentForExcel.UI.ChatView", true);
            var chatView = (System.Windows.FrameworkElement)Activator.CreateInstance(chatViewType);
            var comboField = chatViewType.GetField("QuickModelCombo", BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public);
            var modelCombo = comboField?.GetValue(chatView) as System.Windows.Controls.ComboBox;
            if (modelCombo != null)
            {
                var profiles = settingsType.GetProperty("Profiles").GetValue(settings, null);
                modelCombo.ItemsSource = (System.Collections.IEnumerable)profiles;
                modelCombo.SelectedIndex = 0;
            }
            var messageCollection = chatViewType.GetField("_messages", BindingFlags.Instance | BindingFlags.NonPublic)
                ?.GetValue(chatView) as IList;
            if (messageCollection != null)
            {
                var roleType = _agentAssembly.GetType("AgentForExcel.Models.ChatRole", true);
                var messageType = _agentAssembly.GetType("AgentForExcel.Models.ChatMessage", true);
                var planType = _agentAssembly.GetType("AgentForExcel.Models.TaskPlanData", true);
                var stepType = _agentAssembly.GetType("AgentForExcel.Models.TaskPlanStepData", true);
                var plan = Activator.CreateInstance(planType);
                planType.GetProperty("Title").SetValue(plan, "清洗并生成销售分析", null);
                planType.GetProperty("ProgressText").SetValue(plan, "已完成 1/4 步", null);
                planType.GetProperty("Badge").SetValue(plan, "1/4", null);
                var stepListType = typeof(List<>).MakeGenericType(stepType);
                var stepList = (IList)Activator.CreateInstance(stepListType);
                var statuses = new[] { "completed", "in_progress", "pending", "pending" };
                var titles = new[] { "读取并体检字段", "清洗空值和重复项", "生成分析图表", "回读并验收结果" };
                for (var index = 0; index < titles.Length; index++)
                {
                    var step = Activator.CreateInstance(stepType);
                    stepType.GetProperty("Index").SetValue(step, index + 1, null);
                    stepType.GetProperty("Title").SetValue(step, titles[index], null);
                    stepType.GetProperty("Status").SetValue(step, statuses[index], null);
                    if (index == 0) stepType.GetProperty("Detail").SetValue(step, "字段类型与数据粒度已确认", null);
                    stepList.Add(step);
                }
                planType.GetProperty("Steps").SetValue(plan, stepList, null);
                planType.GetProperty("SuccessCriteria").SetValue(plan,
                    new List<string> { "源数据保持不变", "输出结果可回读核验" }, null);
                var message = Activator.CreateInstance(messageType,
                    new object[] { string.Empty, Enum.Parse(roleType, "Assistant") });
                messageType.GetProperty("Kind").SetValue(message,
                    Enum.Parse(_agentAssembly.GetType("AgentForExcel.Models.ChatMessageKind", true), "TaskPlan"), null);
                messageType.GetProperty("TaskPlan").SetValue(message, plan, null);
                messageCollection.Add(message);
                var welcome = chatViewType.GetField("WelcomePanel", BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public)
                    ?.GetValue(chatView) as System.Windows.UIElement;
                if (welcome != null) welcome.Visibility = System.Windows.Visibility.Collapsed;
            }
            var host = new System.Windows.Window
            {
                Content = chatView,
                Width = 380,
                Height = 660,
                ShowInTaskbar = false,
                WindowStyle = System.Windows.WindowStyle.None,
                Left = -10000,
                Top = -10000
            };
            RenderWindow(host, Path.Combine(directory, "chat-composer-model-switch.png"), 380, 660);
        }

        private static void RenderWindow(System.Windows.Window window, string path, int width, int height)
        {
            try
            {
                window.Width = width;
                window.Height = height;
                window.ShowInTaskbar = false;
                window.WindowStartupLocation = System.Windows.WindowStartupLocation.Manual;
                window.Left = -10000;
                window.Top = -10000;
                window.Show();
                window.UpdateLayout();
                var bitmap = new RenderTargetBitmap(width, height, 96, 96, PixelFormats.Pbgra32);
                bitmap.Render(window);
                var encoder = new PngBitmapEncoder();
                encoder.Frames.Add(BitmapFrame.Create(bitmap));
                using (var stream = File.Create(path)) encoder.Save(stream);
            }
            finally
            {
                window.Close();
            }
        }

        private static void Assert(bool condition, string message)
        {
            if (!condition) throw new InvalidOperationException(message);
        }

        private static void ReleaseCom(object value)
        {
            try
            {
                if (value != null && Marshal.IsComObject(value)) Marshal.FinalReleaseComObject(value);
            }
            catch { }
        }
    }
}
