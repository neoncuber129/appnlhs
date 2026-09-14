using System.IO;
using Microsoft.Win32;

namespace ExcelDataEntryApp.Services;

/// <summary>Kiểm tra Excel desktop có cài hay không mà không load assembly Office.Interop.</summary>
public static class ExcelInstallationProbe
{
    private static bool? _installed;

    public static bool IsInstalled()
    {
        if (!OperatingSystem.IsWindows())
        {
            return false;
        }

        if (_installed.HasValue)
        {
            return _installed.Value;
        }

        try
        {
            _installed = HasExcelExecutable() && Type.GetTypeFromProgID("Excel.Application") is not null;
        }
        catch
        {
            _installed = false;
        }

        return _installed.Value;
    }

    private static bool HasExcelExecutable()
    {
        const string appPathsKey = @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe";
        foreach (var view in new[] { RegistryView.Registry64, RegistryView.Registry32 })
        {
            using var baseKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, view);
            using var excelKey = baseKey.OpenSubKey(appPathsKey);
            var path = excelKey?.GetValue(null) as string;
            if (!string.IsNullOrWhiteSpace(path) && File.Exists(path))
            {
                return true;
            }
        }

        return false;
    }

    public static void MarkUnavailable()
    {
        _installed = false;
    }
}
