# 性能日志说明

加载项会把本地诊断和性能日志写入：`%LOCALAPPDATA%\AgentForExcel\startup.log`。

性能记录以 `PERF|` 开头，只保存组件、耗时、操作名和数据规模等技术指标；不会保存单元格内容、工作簿路径、提示词或 API Key。

常见记录：

- `component=operation`：每次 Agent 工具调用的总耗时。
- `component=range_read`：读取区域的尺寸与实际预览尺寸。
- `component=data_profile`：数据体检的行列数、单元格数与耗时。
- `component=chat_run`：一次对话的总耗时、轮次、工具调用数和流式渲染次数。

在 PowerShell 中查看最近 300 条性能记录：

```powershell
Get-Content "$env:LOCALAPPDATA\AgentForExcel\startup.log" -Tail 300 |
  Select-String 'PERF\|'
```

优化时优先对比同一操作、相近数据规模下的 `elapsed_ms`；不要把不同模型响应时间与本地 Excel 操作耗时混为一类。
