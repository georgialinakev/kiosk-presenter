using static KioskPresenter.Host.Interop.NativeMethods;

namespace KioskPresenter.Host.Interop;

internal static class WindowHelper
{
    /// <summary>
    /// Brings a window to the front even when Windows' focus-stealing rules would refuse a plain
    /// SetForegroundWindow (another process holds the foreground). Attaches to the foreground
    /// thread's input for the duration, which is the documented way to share activation.
    /// </summary>
    public static void ForceForeground(IntPtr hwnd)
    {
        if (hwnd == IntPtr.Zero || !IsWindow(hwnd)) return;
        if (IsIconic(hwnd)) ShowWindow(hwnd, SW_RESTORE);

        var foreground = GetForegroundWindow();
        if (foreground == hwnd) return;

        var foregroundThread = GetWindowThreadProcessId(foreground, out _);
        var ourThread = GetCurrentThreadId();
        var attached = foregroundThread != 0 && foregroundThread != ourThread && AttachThreadInput(ourThread, foregroundThread, true);
        try
        {
            BringWindowToTop(hwnd);
            ShowWindow(hwnd, SW_SHOW);
            if (!SetForegroundWindow(hwnd))
            {
                // An Alt tap counts as user input, which lifts the foreground lock for this call.
                const byte VK_MENU = 0x12;
                const uint KEYEVENTF_KEYUP = 0x0002;
                keybd_event(VK_MENU, 0, 0, UIntPtr.Zero);
                keybd_event(VK_MENU, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
                SetForegroundWindow(hwnd);
            }
        }
        finally
        {
            if (attached) AttachThreadInput(ourThread, foregroundThread, false);
        }
    }

    public static void SetTopmost(IntPtr hwnd, bool topmost) =>
        SetWindowPos(hwnd, topmost ? HWND_TOPMOST : HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);

    /// <summary>Makes a window that never takes focus and never shows in Alt+Tab.</summary>
    public static void MakeNoActivate(IntPtr hwnd)
    {
        var style = GetWindowLongPtr(hwnd, GWL_EXSTYLE).ToInt64();
        style |= WS_EX_NOACTIVATE | WS_EX_TOOLWINDOW;
        SetWindowLongPtr(hwnd, GWL_EXSTYLE, new IntPtr(style));
    }

    /// <summary>Keeps the display on while the kiosk runs; no power-plan changes needed.</summary>
    public static void KeepDisplayOn(bool on) =>
        SetThreadExecutionState(on ? ES_CONTINUOUS | ES_DISPLAY_REQUIRED | ES_SYSTEM_REQUIRED : ES_CONTINUOUS);
}
