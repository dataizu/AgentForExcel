namespace AgentForExcel.Operations
{
    /// <summary>
    /// 集中注册所有操作工厂。随阶段推进逐步解锁各能力。
    /// 启动时由 AppContext.Initialize() 调用一次。
    /// </summary>
    public static class OperationRegistration
    {
        public static void RegisterAll(OperationDispatcher dispatcher, AppContext context)
        {
            // ---- 环境诊断 ----
            dispatcher.Register(new SystemCheck.AgentSelfCheckOp.Factory());

            // ---- 任务执行闭环：计划、步骤状态与完成检查 ----
            dispatcher.Register(new Tasking.PlanTaskOp.Factory());
            dispatcher.Register(new Tasking.UpdateTaskStepOp.Factory());

            // ---- 阶段 1:只读工具(联调用,可直接放行)----
            dispatcher.Register(new Cell.ReadRangeOp.Factory());
            dispatcher.Register(new Analysis.ProfileDataOp.Factory());

            // ---- 阶段 2:单元格执行闭环(写操作统一经过确认)----
            dispatcher.Register(new Cell.WriteRangeOp.Factory());
            dispatcher.Register(new Cell.FillFormulaOp.Factory());
            dispatcher.Register(new Cell.FormatRangeOp.Factory());
            dispatcher.Register(new Cell.DrawPixelsOp.Factory());
            dispatcher.Register(new Cell.DrawFromImageOp.Factory());
            dispatcher.Register(new Analysis.CreateAnalysisViewOp.Factory());

            // ---- 阶段 3:图表与普通数据透视表 ----
            dispatcher.Register(new Chart.CreateChartOp.Factory());
            dispatcher.Register(new Pivot.CreatePivotTableOp.Factory());
            dispatcher.Register(new Dashboard.CreateDashboardOp.Factory());

            // ---- 阶段 4:Power Query 数据清洗与可刷新加载 ----
            dispatcher.Register(new PowerQuery.ListQueriesOp.Factory());
            dispatcher.Register(new PowerQuery.CreateRangeQueryOp.Factory());
            dispatcher.Register(new PowerQuery.LoadQueryOp.Factory());
            dispatcher.Register(new PowerQuery.RefreshQueryOp.Factory());

            // ---- 阶段 5:Power Pivot 数据模型、关系、DAX 与模型透视 ----
            dispatcher.Register(new PowerPivot.ListModelOp.Factory());
            dispatcher.Register(new PowerPivot.AddQueryToModelOp.Factory());
            dispatcher.Register(new PowerPivot.RefreshModelOp.Factory());
            dispatcher.Register(new PowerPivot.AddRelationshipOp.Factory());
            dispatcher.Register(new PowerPivot.AddMeasureOp.Factory());
            dispatcher.Register(new PowerPivot.CreateModelPivotOp.Factory());

            // ---- 阶段 6:受控 VBA（白名单预览、显式确认、备份、临时执行与审计）----
            dispatcher.Register(new Macro.PreviewSafeVbaOp.Factory());
            dispatcher.Register(new Macro.ExecuteSafeVbaOp.Factory());
        }
    }
}
