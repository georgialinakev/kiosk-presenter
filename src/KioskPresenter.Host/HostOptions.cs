using KioskPresenter.Core;

namespace KioskPresenter.Host;

/// <summary>
/// Command line. A kiosk starts with no arguments; the switches are for testing on a workstation.
/// </summary>
/// <remarks>
/// --windowed      Normal window, no keyboard hook, no focus guard, no watchdog. For development.
/// --devtools      Enables WebView2 dev tools and the context menu.
/// --no-watchdog   Kiosk mode without the watchdog.
/// --data PATH     Use PATH instead of C:\ProgramData\KioskPresenter.
/// --watchdog      Internal: run as the watchdog for the host process given by --pid.
/// </remarks>
internal sealed record HostOptions(bool Windowed, bool DevTools, bool NoWatchdog, bool WatchdogMode, int WatchPid, string? DataRoot)
{
    public static HostOptions Parse(string[] args)
    {
        bool Has(string flag) => args.Any(a => string.Equals(a, flag, StringComparison.OrdinalIgnoreCase));
        string? Value(string flag)
        {
            var i = Array.FindIndex(args, a => string.Equals(a, flag, StringComparison.OrdinalIgnoreCase));
            return i >= 0 && i + 1 < args.Length ? args[i + 1] : null;
        }

        var pid = int.TryParse(Value("--pid"), out var p) ? p : 0;
        return new HostOptions(Has("--windowed"), Has("--devtools"), Has("--no-watchdog"), Has("--watchdog"), pid, Value("--data"));
    }

    public KioskPaths Paths => DataRoot is null ? KioskPaths.Default : new KioskPaths(DataRoot);

    /// <summary>Hooks, focus guard and topmost only make sense on a real kiosk.</summary>
    public bool KioskMode => !Windowed;

    public bool UseWatchdog => KioskMode && !NoWatchdog;
}
