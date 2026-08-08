using System;
using System.Text.RegularExpressions;

namespace AgentForExcel.Services
{
    /// <summary>
    /// 记录可用于性能优化的本地指标。日志只允许组件名、耗时和已白名单化的
    /// 数字/状态元数据，避免把单元格内容、工作簿路径、提示词或 API Key 写入磁盘。
    /// </summary>
    internal static class PerformanceLogger
    {
        public static void Log(string component, long elapsedMilliseconds, string metadata = null)
        {
            var entry = "PERF|component=" + Sanitize(component) +
                        "|elapsed_ms=" + Math.Max(0, elapsedMilliseconds);
            if (!string.IsNullOrWhiteSpace(metadata))
                entry += "|" + Sanitize(metadata);
            ThisAddIn.Log(entry);
        }

        private static string Sanitize(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return "unknown";
            // 仅保留结构化性能字段所需的字符；任何意外传入的业务内容都不会原样落盘。
            return Regex.Replace(value, "[^A-Za-z0-9_.=|:-]", "_");
        }
    }
}
