using System.Collections.Generic;

namespace AgentForExcel.AI
{
    /// <summary>
    /// Function Calling 工具定义(供 OpenAI 兼容协议的 tools 字段使用)。
    ///
    /// 设计原则:
    /// - 每个能力(单元格 / 图表 / Power Query / Power Pivot / 宏)对应若干工具。
    /// - 工具名前缀即能力域(如 ""cell.write_range""),便于派发器按前缀路由。
    /// - 阶段 2 起逐步解锁各工具;当前阶段 1 先放一个只读工具供联调。
    /// </summary>
    public static class ToolDefinitions
    {
        /// <summary>随请求一并发给 LLM 的工具清单。</summary>
        public static readonly IReadOnlyList<object> Tools = BuildTools();

        private static List<object> BuildTools()
        {
            var tools = new List<object>();
            tools.Add(MakeTool(
                name: "agent_self_check",
                description: "只读检查当前 Excel、工作簿、Power Query、Power Pivot 和受控 VBA 权限状态。用户反馈能力不可用或正式执行前可先调用。",
                props: new object[0],
                required: new string[0]));

            tools.Add(MakeTool(
                name: "task_plan",
                description: "为需要多个步骤、写入工作簿或生成交付物的任务建立可见执行计划。计划创建后必须用 task_step_update 持续更新，所有步骤完成前不得结束任务。",
                props: new object[]
                {
                    new { name = "title", type = "string", description = "简洁的任务标题" },
                    new { name = "steps", type = "array", description = "按执行顺序排列的具体步骤，建议 2 到 8 项", items = new { type = "string" } },
                    new { name = "success_criteria", type = "array", description = "完成任务必须满足的验收条件", items = new { type = "string" } }
                },
                required: new[] { "steps" }));

            tools.Add(MakeTool(
                name: "task_step_update",
                description: "更新任务计划中某一步的状态。只有实际执行并核验后才能标记 completed；失败时使用 failed 并说明原因。",
                props: new object[]
                {
                    new { name = "step_index", type = "integer", description = "计划中的步骤序号，从 1 开始" },
                    new { name = "status", type = "string", description = "步骤状态", @enum = new[] { "pending", "in_progress", "completed", "failed" } },
                    new { name = "detail", type = "string", description = "执行结果、核验证据或失败原因" }
                },
                required: new[] { "step_index", "status" }));

            // ---- 阶段 1:联调用,只读工具(无副作用,可直接放行)----
            tools.Add(MakeTool(
                name: "cell_read_range",
                description: "读取工作表中指定区域的值。只读,无副作用。",
                props: new object[]
                {
                    new { name = "sheet", type = "string", description = "工作表名(可选,默认为活动工作表)" },
                    new { name = "address", type = "string", description = "区域地址,如 \"A1:B10\"。省略时读取已用区域" }
                },
                required: new[] { "address" }));

            tools.Add(MakeTool(
                name: "data_profile",
                description: "在分析、制图或创建看板前，对二维数据做只读体检：识别字段类型和角色、缺失率、不同值数量、重复记录、零值和 IQR 异常值，并给出可派生时间维度与图表基数风险。不会修改工作簿。",
                props: new object[]
                {
                    new { name = "sheet", type = "string", description = "工作表名(可选,默认为活动工作表)" },
                    new { name = "address", type = "string", description = "包含唯一非空标题行的完整分析区域,如 A1:I345" }
                },
                required: new[] { "address" }));

            // ---- 阶段 2:安全的单元格写入、公式与格式 ----
            tools.Add(MakeTool(
                name: "cell_write_range",
                description: "向工作表区域写入普通值。values 支持二维数组；address 为单个起始单元格时会按数组尺寸扩展。禁止用该工具写公式。",
                props: new object[]
                {
                    new { name = "sheet", type = "string", description = "工作表名(可选,默认为活动工作表)" },
                    new { name = "address", type = "string", description = "目标区域或左上角单元格,如 A2:C10 或 A2" },
                    new
                    {
                        name = "values",
                        type = "array",
                        description = "要写入的二维行列数组,例如 [[\"产品\",\"销售额\"],[\"A\",100]]",
                        items = new { type = "array", items = new { } }
                    }
                },
                required: new[] { "address", "values" }));

            tools.Add(MakeTool(
                name: "cell_fill_formula",
                description: "向指定区域填充工作簿内计算公式。默认使用 A1 公式；批量相对引用建议使用 R1C1。禁止外部调用类公式。",
                props: new object[]
                {
                    new { name = "sheet", type = "string", description = "工作表名(可选,默认为活动工作表)" },
                    new { name = "address", type = "string", description = "目标区域,如 D2:D100" },
                    new { name = "formula", type = "string", description = "以 = 开头的 Excel 公式" },
                    new { name = "use_r1c1", type = "boolean", description = "是否使用 R1C1 记法,默认 false" }
                },
                required: new[] { "address", "formula" }));

            tools.Add(MakeTool(
                name: "cell_format_range",
                description: "设置区域的数字格式、字体、填充、对齐、边框、列宽或行高。只传需要修改的字段。颜色必须为 #RRGGBB。",
                props: new object[]
                {
                    new { name = "sheet", type = "string", description = "工作表名(可选,默认为活动工作表)" },
                    new { name = "address", type = "string", description = "目标区域,如 A1:F20" },
                    new { name = "number_format", type = "string", description = "Excel 数字格式,如 0.00、0%、yyyy-mm-dd" },
                    new { name = "bold", type = "boolean", description = "是否加粗" },
                    new { name = "italic", type = "boolean", description = "是否斜体" },
                    new { name = "font_size", type = "number", description = "字号,范围 6 到 72" },
                    new { name = "font_color", type = "string", description = "字体颜色,#RRGGBB" },
                    new { name = "fill_color", type = "string", description = "填充颜色,#RRGGBB" },
                    new { name = "horizontal_alignment", type = "string", description = "水平对齐", @enum = new[] { "left", "center", "right", "general" } },
                    new { name = "wrap_text", type = "boolean", description = "是否自动换行" },
                    new { name = "add_borders", type = "boolean", description = "是否添加细边框" },
                    new { name = "autofit_columns", type = "boolean", description = "是否自适应列宽" },
                    new { name = "autofit_rows", type = "boolean", description = "是否自适应行高" }
                },
                required: new[] { "address" }));

            tools.Add(MakeTool(
                name: "analysis_create_view",
                description: "把源区域复制为值快照，在新的分析工作表中安全排序和展示。不会修改源数据，也不会把源公式带入分析页。用户要求排序、筛选式展示或分析副本时优先使用。",
                props: new object[]
                {
                    new { name = "source_sheet", type = "string", description = "数据源工作表名(可选,默认为活动工作表)" },
                    new { name = "source_address", type = "string", description = "包含唯一非空标题行的数据源区域,如 A1:F100" },
                    new { name = "analysis_sheet_name", type = "string", description = "新分析工作表的名称前缀,默认 Agent分析；重名时自动添加序号" },
                    new { name = "destination_address", type = "string", description = "分析表中快照左上角,默认 A1" },
                    new
                    {
                        name = "sort_by",
                        type = "array",
                        description = "排序字段,最多 3 个；只对新分析页的快照排序",
                        items = new
                        {
                            type = "object",
                            properties = new
                            {
                                field = new { type = "string", description = "源区域中的字段名" },
                                direction = new { type = "string", @enum = new[] { "asc", "desc" }, description = "升序或降序" }
                            },
                            required = new[] { "field" }
                        }
                    }
                },
                required: new[] { "source_address" }));

            // ---- 阶段 3:图表与普通数据透视表 ----
            tools.Add(MakeTool(
                name: "chart_create",
                description: "根据二维表区域创建可直接用于报告的现代图表。默认新建 Agent图表 工作表，保留值快照且不修改源数据。条形、折线、面积、散点和占比图会自动应用对应的专业样式；占比场景优先用 doughnut，并明确 category_field 与 value_field。",
                props: new object[]
                {
                    new { name = "source_sheet", type = "string", description = "数据源工作表名(可选,默认为活动工作表)" },
                    new { name = "source_address", type = "string", description = "包含标题行的数据源区域,如 A1:D20" },
                    new { name = "destination_sheet", type = "string", description = "图表放置工作表(可选)。不传时自动创建新的 Agent图表 分析页并复制值快照" },
                    new { name = "anchor_address", type = "string", description = "图表左上角锚点单元格(可选,新分析页会自动避开数据表)" },
                    new { name = "chart_type", type = "string", description = "图表类型；占比分析优先使用 doughnut", @enum = new[] { "column", "bar", "line", "pie", "doughnut", "area", "scatter" } },
                    new { name = "title", type = "string", description = "图表标题" },
                    new { name = "name", type = "string", description = "Excel 图表对象名称(可选)" },
                    new { name = "category_field", type = "string", description = "分类字段名。占比图必须传入,如 Category 或 类别" },
                    new { name = "value_field", type = "string", description = "数值字段名。占比图必须传入,如 Actual 或 实际支出，避免选错指标" },
                    new { name = "aggregation", type = "string", description = "同一分类出现多行时的聚合方式。auto 会在发现重复分类时自动求和", @enum = new[] { "auto", "none", "sum", "average", "count", "min", "max" } },
                    new { name = "max_categories", type = "integer", description = "分类图允许的最大可见分类数,范围 3 到 50；普通分类图默认 12，占比图默认 8" },
                    new { name = "include_other", type = "boolean", description = "分类超过上限时是否把长尾汇总为“其他”；占比图默认 true" },
                    new { name = "exclude_categories", type = "array", description = "不参与图表的分类值,如 [\"储蓄\",\"总计\"]", items = new { type = "string" } },
                    new { name = "sort_descending", type = "boolean", description = "是否按数值从高到低排列；占比图默认 true，只改变图表快照顺序" },
                    new { name = "show_data_labels", type = "boolean", description = "是否显示数据标签；柱形、条形和占比图默认 true，折线、面积和散点图按需要开启" },
                    new { name = "show_percentage", type = "boolean", description = "数据标签是否显示百分比；占比图默认 true" },
                    new { name = "legend_position", type = "string", description = "图例位置", @enum = new[] { "none", "bottom", "right" } },
                    new { name = "palette", type = "string", description = "协调配色方案", @enum = new[] { "emerald", "ocean", "sunset", "vivid" } },
                    new { name = "width", type = "number", description = "图表宽度,默认 600" },
                    new { name = "height", type = "number", description = "图表高度,默认 380" }
                },
                required: new[] { "source_address", "chart_type" }));

            tools.Add(MakeTool(
                name: "pivot_create",
                description: "根据标准二维表创建普通非 OLAP 数据透视表。源区域首行必须是唯一且非空的字段名。目标工作表不存在时会新建。",
                props: new object[]
                {
                    new { name = "source_sheet", type = "string", description = "数据源工作表名(可选,默认为活动工作表)" },
                    new { name = "source_address", type = "string", description = "包含标题行的数据源区域,如 A1:D100" },
                    new { name = "destination_sheet", type = "string", description = "目标工作表名(可选,默认透视分析;不存在时新建)" },
                    new { name = "destination_address", type = "string", description = "透视表左上角,默认 A1" },
                    new { name = "name", type = "string", description = "透视表名称,默认 AgentPivot" },
                    new { name = "rows", type = "array", description = "行字段名数组", items = new { type = "string" } },
                    new { name = "columns", type = "array", description = "列字段名数组", items = new { type = "string" } },
                    new { name = "filters", type = "array", description = "筛选字段名数组", items = new { type = "string" } },
                    new
                    {
                        name = "values",
                        type = "array",
                        description = "值字段配置数组",
                        items = new
                        {
                            type = "object",
                            properties = new
                            {
                                field = new { type = "string", description = "值字段名" },
                                function = new { type = "string", @enum = new[] { "sum", "count", "average", "max", "min" } },
                                caption = new { type = "string", description = "显示名称(可选)" }
                            },
                            required = new[] { "field" }
                        }
                    }
                },
                required: new[] { "source_address", "values" }));

            tools.Add(MakeTool(
                name: "dashboard_create",
                description: "根据标准二维表创建原生 Excel 联动数据看板。生成 3 个 KPI、动态透视图、明细透视表和最多 3 个全局筛选器。默认优先使用原生切片器，失败时自动改用下拉兼容模式；高级联动由加载项执行，不注入 VBA。源工作表不会被修改。",
                props: new object[]
                {
                    new { name = "source_sheet", type = "string", description = "数据源工作表名(可选,默认为活动工作表)" },
                    new { name = "source_address", type = "string", description = "包含唯一非空标题行的数据源区域,如 A1:H1000" },
                    new { name = "dashboard_sheet_name", type = "string", description = "新看板工作表名称前缀,默认 Agent看板；重名时自动添加序号" },
                    new { name = "title", type = "string", description = "看板标题,默认业务分析看板" },
                    new { name = "date_field", type = "string", description = "日期或时间字段名(可选)。传入后用于趋势图横轴" },
                    new { name = "category_field", type = "string", description = "核心分类字段名，用于排名和明细，如 产品、类别" },
                    new { name = "series_field", type = "string", description = "第二分析维度(可选)，用于占比和明细，如 区域、渠道" },
                    new { name = "value_field", type = "string", description = "核心数值字段名，如 销售额、利润" },
                    new { name = "filter_fields", type = "array", description = "全局联动筛选字段，最多 3 个，如 [\"区域\",\"类别\"]", items = new { type = "string" } },
                    new { name = "filter_mode", type = "string", description = "筛选方式。auto 默认优先切片器并在不兼容时回退下拉；slicer 强制原生切片器；dropdown 使用下拉兼容模式", @enum = new[] { "auto", "slicer", "dropdown" } },
                    new { name = "aggregation", type = "string", description = "图表和明细的聚合方式,默认 sum", @enum = new[] { "sum", "count", "average" } },
                    new { name = "top_n", type = "integer", description = "排名图展示前 N 项,范围 3 到 20,默认 10" },
                    new { name = "number_format", type = "string", description = "指标数字格式,如 ¥#,##0、0.0% 或 #,##0.00" }
                },
                required: new[] { "source_address", "category_field", "value_field" }));

            tools.Add(MakeTool(
                name: "pq_list_queries",
                description: "列出当前工作簿中的 Power Query 查询名称、说明和 M 公式预览。只读。",
                props: new object[0],
                required: new string[0]));

            tools.Add(MakeTool(
                name: "pq_create_from_range",
                description: "根据当前工作簿中的二维区域创建结构化 Power Query 清洗查询。使用隐藏名称引用源区域，不会修改源数据；支持删除空行、清理文本、去重、重命名、字段类型和选列。",
                props: new object[]
                {
                    new { name = "source_sheet", type = "string", description = "源工作表名(可选,默认为活动工作表)" },
                    new { name = "source_address", type = "string", description = "包含标题行的完整源区域,如 A1:H1000" },
                    new { name = "query_name", type = "string", description = "Power Query 查询名称" },
                    new { name = "remove_blank_rows", type = "boolean", description = "是否删除完全空白行,默认 true" },
                    new { name = "trim_text", type = "boolean", description = "是否清理文本字段首尾空格,默认 true" },
                    new { name = "remove_duplicates", type = "boolean", description = "是否删除完全重复记录,默认 true" },
                    new { name = "replace_existing", type = "boolean", description = "同名查询存在时是否更新,默认 false" },
                    new
                    {
                        name = "rename_columns",
                        type = "array",
                        description = "字段重命名规则",
                        items = new
                        {
                            type = "object",
                            properties = new
                            {
                                from = new { type = "string" },
                                to = new { type = "string" }
                            },
                            required = new[] { "from", "to" }
                        }
                    },
                    new
                    {
                        name = "column_types",
                        type = "array",
                        description = "字段类型规则",
                        items = new
                        {
                            type = "object",
                            properties = new
                            {
                                field = new { type = "string" },
                                type = new
                                {
                                    type = "string",
                                    @enum = new[] { "text", "number", "integer", "date", "datetime", "logical", "currency", "percentage" }
                                }
                            },
                            required = new[] { "field", "type" }
                        }
                    },
                    new { name = "select_columns", type = "array", description = "最终保留的字段名列表(可选)", items = new { type = "string" } }
                },
                required: new[] { "source_address", "query_name" }));

            tools.Add(MakeTool(
                name: "pq_load_to_sheet",
                description: "把已有 Power Query 查询同步加载为新工作表中的 Excel 表格，并返回实际行列数。默认不覆盖现有工作表。",
                props: new object[]
                {
                    new { name = "query_name", type = "string", description = "要加载的查询名称" },
                    new { name = "destination_sheet", type = "string", description = "新结果工作表名称前缀,默认 Agent查询结果" },
                    new { name = "destination_address", type = "string", description = "加载起始单元格,默认 A1" }
                },
                required: new[] { "query_name" }));

            tools.Add(MakeTool(
                name: "pq_refresh",
                description: "刷新指定 Power Query 查询已经加载到工作表的全部结果。",
                props: new object[]
                {
                    new { name = "query_name", type = "string", description = "要刷新的查询名称" }
                },
                required: new[] { "query_name" }));

            tools.Add(MakeTool(
                name: "pp_list_model",
                description: "读取当前工作簿 Power Pivot 数据模型中的表、字段、关系和 DAX 度量值。只读。",
                props: new object[0],
                required: new string[0]));

            tools.Add(MakeTool(
                name: "pp_add_query_to_model",
                description: "把已经创建的 Power Query 查询载入 Power Pivot 数据模型。先用 pq_create_from_range 创建查询。",
                props: new object[]
                {
                    new { name = "query_name", type = "string", description = "要载入数据模型的 Power Query 查询名称" }
                },
                required: new[] { "query_name" }));

            tools.Add(MakeTool(
                name: "pp_refresh_model",
                description: "刷新 Power Pivot 数据模型及基于模型创建的透视表。源数据或 Power Query 结果变化后使用。",
                props: new object[0],
                required: new string[0]));

            tools.Add(MakeTool(
                name: "pp_add_relationship",
                description: "在 Power Pivot 数据模型中建立一对多关系。from 是多端外键，to 是一端唯一键。",
                props: new object[]
                {
                    new { name = "from_table", type = "string", description = "多端模型表" },
                    new { name = "from_column", type = "string", description = "多端外键字段" },
                    new { name = "to_table", type = "string", description = "一端模型表" },
                    new { name = "to_column", type = "string", description = "一端唯一键字段" }
                },
                required: new[] { "from_table", "from_column", "to_table", "to_column" }));

            tools.Add(MakeTool(
                name: "pp_add_measure",
                description: "在 Power Pivot 模型表中创建 DAX 度量值。formula 只写表达式，例如 SUM(Sales[Amount])，不要写 EVALUATE/DEFINE。",
                props: new object[]
                {
                    new { name = "table", type = "string", description = "度量值所属模型表" },
                    new { name = "measure_name", type = "string", description = "度量值名称" },
                    new { name = "formula", type = "string", description = "DAX 表达式，可带或不带开头等号" },
                    new { name = "format", type = "string", @enum = new[] { "general", "currency", "whole_number", "decimal", "percentage", "date", "scientific", "boolean" }, description = "显示格式，默认 general" },
                    new { name = "decimal_places", type = "integer", description = "小数位数，默认 2" },
                    new { name = "currency_symbol", type = "string", description = "币种符号，例如 ¥ 或 $" },
                    new { name = "description", type = "string", description = "业务口径说明" },
                    new { name = "replace_existing", type = "boolean", description = "是否替换同名度量值，默认 false" }
                },
                required: new[] { "table", "measure_name", "formula" }));

            var modelFieldReference = new
            {
                type = "object",
                properties = new
                {
                    table = new { type = "string", description = "模型表名" },
                    field = new { type = "string", description = "字段名" }
                },
                required = new[] { "table", "field" }
            };
            tools.Add(MakeTool(
                name: "pp_create_model_pivot",
                description: "基于 Power Pivot 数据模型创建可直接使用的模型透视表，可跨表使用维度并展示 DAX 度量值。",
                props: new object[]
                {
                    new { name = "destination_sheet", type = "string", description = "新工作表名称前缀，默认 Agent模型透视" },
                    new { name = "destination_address", type = "string", description = "起始单元格，默认 A1" },
                    new { name = "name", type = "string", description = "透视表名称，默认 AgentModelPivot" },
                    new { name = "rows", type = "array", description = "行维度", items = modelFieldReference },
                    new { name = "columns", type = "array", description = "列维度", items = modelFieldReference },
                    new { name = "filters", type = "array", description = "筛选维度", items = modelFieldReference },
                    new { name = "measures", type = "array", description = "要显示的 DAX 度量值名称", items = new { type = "string" } }
                },
                required: new[] { "measures" }));

            tools.Add(MakeTool(
                name: "vba_preview_safe",
                description: "预览受控 VBA 白名单配方，只生成代码、风险说明和一次性令牌，不修改工作簿。执行前必须先调用此工具。",
                props: new object[]
                {
                    new { name = "recipe", type = "string", description = "白名单配方：refresh_all、autofit_used_ranges、export_active_sheet_pdf" },
                    new { name = "output_path", type = "string", description = "仅导出 PDF 时可选；必须是本地 .pdf 完整路径" }
                },
                required: new[] { "recipe" }));

            tools.Add(MakeTool(
                name: "vba_execute_safe",
                description: "执行已经预览的受控 VBA。需要 vba_preview_safe 返回的一次性令牌；执行前会触发用户确认并创建备份，结束后删除临时模块并写审计日志。",
                props: new object[]
                {
                    new { name = "preview_token", type = "string", description = "vba_preview_safe 返回的 preview_token" }
                },
                required: new[] { "preview_token" }));

            return tools;
        }

        /// <summary>构造一个工具定义对象(OpenAI function calling 格式)。</summary>
        private static object MakeTool(string name, string description, object[] props, string[] required)
        {
            return new
            {
                type = "function",
                function = new
                {
                    name,
                    description,
                    parameters = new
                    {
                        type = "object",
                        properties = ToPropertyDict(props),
                        required
                    }
                }
            };
        }

        /// <summary>把匿名属性数组转成字典(JSON 序列化后即为 JSON 对象)。</summary>
        private static Dictionary<string, object> ToPropertyDict(object[] props)
        {
            var dict = new Dictionary<string, object>();
            if (props == null) return dict;
            foreach (var p in props)
            {
                // 用反射取出匿名对象的 name 字段作为 key
                var name = p.GetType().GetProperty("name")?.GetValue(p)?.ToString();
                if (name != null) dict[name] = p;
            }
            return dict;
        }
    }
}
