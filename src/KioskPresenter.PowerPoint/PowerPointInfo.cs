using System.Diagnostics;
using System.IO;
using Microsoft.Win32;

namespace KioskPresenter.PowerPoint;

/// <summary>Finds the installed PowerPoint without starting it.</summary>
public static class PowerPointInfo
{
    public const string ProgId = "PowerPoint.Application";

    /// <summary>"Microsoft PowerPoint 16.0.17928.20114", or null when PowerPoint isn't installed.</summary>
    public static string? Describe()
    {
        if (Type.GetTypeFromProgID(ProgId) is null) return null;
        var exe = ExePath();
        if (exe is null || !File.Exists(exe)) return "Microsoft PowerPoint";
        try
        {
            var version = FileVersionInfo.GetVersionInfo(exe);
            return $"Microsoft PowerPoint {version.FileVersion}";
        }
        catch (Exception)
        {
            return "Microsoft PowerPoint";
        }
    }

    public static bool IsInstalled => Type.GetTypeFromProgID(ProgId) is not null;

    /// <summary>POWERPNT.EXE from App Paths (Click-to-Run and MSI installs both register it).</summary>
    public static string? ExePath()
    {
        const string key = @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\powerpnt.exe";
        foreach (var view in new[] { RegistryView.Registry64, RegistryView.Registry32 })
        {
            try
            {
                using var root = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, view);
                using var sub = root.OpenSubKey(key);
                if (sub?.GetValue(null) is string path && path.Length > 0) return path.Trim('"');
            }
            catch (Exception)
            {
                // Try the other view.
            }
        }
        return null;
    }
}
