using System;
using Microsoft.Win32;

namespace AgentForExcel.Services
{
    public sealed class RuntimeEnvironmentSnapshot
    {
        public string OperatingSystem { get; set; }
        public bool Is64BitOperatingSystem { get; set; }
        public bool Is64BitExcelProcess { get; set; }
        public int DotNetRelease { get; set; }
        public bool DotNet48OrLater { get; set; }
        public string VstoRuntimeVersion { get; set; }
        public bool VstoRuntimeDetected { get; set; }
    }

    /// <summary>Read-only install diagnostics for support and first-run checks.</summary>
    public static class RuntimeEnvironmentInspector
    {
        private const int DotNet48Release = 528040;

        public static RuntimeEnvironmentSnapshot Capture()
        {
            var snapshot = new RuntimeEnvironmentSnapshot
            {
                OperatingSystem = Environment.OSVersion.VersionString,
                Is64BitOperatingSystem = Environment.Is64BitOperatingSystem,
                Is64BitExcelProcess = Environment.Is64BitProcess
            };

            snapshot.DotNetRelease = ReadDword(@"SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full", "Release");
            snapshot.DotNet48OrLater = snapshot.DotNetRelease >= DotNet48Release;
            snapshot.VstoRuntimeVersion = ReadFirstString(
                @"SOFTWARE\Microsoft\VSTO Runtime Setup\v4R",
                @"SOFTWARE\Microsoft\VSTO Runtime Setup\v4");
            snapshot.VstoRuntimeDetected = !string.IsNullOrWhiteSpace(snapshot.VstoRuntimeVersion);
            return snapshot;
        }

        private static int ReadDword(string path, string valueName)
        {
            foreach (var view in GetViews())
            {
                try
                {
                    using (var baseKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, view))
                    using (var key = baseKey.OpenSubKey(path))
                    {
                        var value = key?.GetValue(valueName);
                        if (value != null) return Convert.ToInt32(value);
                    }
                }
                catch { }
            }
            return 0;
        }

        private static string ReadFirstString(params string[] paths)
        {
            foreach (var path in paths)
            {
                foreach (var view in GetViews())
                {
                    try
                    {
                        using (var baseKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, view))
                        using (var key = baseKey.OpenSubKey(path))
                        {
                            var value = key?.GetValue("Version") as string;
                            if (!string.IsNullOrWhiteSpace(value)) return value;
                        }
                    }
                    catch { }
                }
            }
            return null;
        }

        private static RegistryView[] GetViews()
        {
            return Environment.Is64BitOperatingSystem
                ? new[] { RegistryView.Registry64, RegistryView.Registry32 }
                : new[] { RegistryView.Registry32 };
        }
    }
}
