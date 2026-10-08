using static KioskPresenter.Host.Interop.NativeMethods;

namespace KioskPresenter.Host.Interop;

/// <summary>A window from another process just became the foreground window.</summary>
internal sealed record ForegroundChange(IntPtr Hwnd, int ProcessId, string ClassName);

/// <summary>
/// Reports every foreground change in the session. The controller decides whether it's allowed
/// (our windows, the PowerPoint slideshow) or must be undone.
/// </summary>
internal sealed class ForegroundWatcher : IDisposable
{
    private readonly WinEventDelegate _proc;
    private IntPtr _hook;

    public ForegroundWatcher() => _proc = Callback;

    public event Action<ForegroundChange>? Changed;

    public void Start()
    {
        if (_hook != IntPtr.Zero) return;
        // Out-of-context: Windows posts the event to this (UI) thread's message queue.
        _hook = SetWinEventHook(EVENT_SYSTEM_FOREGROUND, EVENT_SYSTEM_FOREGROUND, IntPtr.Zero, _proc, 0, 0, WINEVENT_OUTOFCONTEXT);
        if (_hook == IntPtr.Zero) Core.Log.Error("Foreground hook failed.");
    }

    public void Stop()
    {
        if (_hook == IntPtr.Zero) return;
        UnhookWinEvent(_hook);
        _hook = IntPtr.Zero;
    }

    private void Callback(IntPtr hook, uint eventType, IntPtr hwnd, int idObject, int idChild, uint thread, uint time)
    {
        if (hwnd == IntPtr.Zero) return;
        try
        {
            Changed?.Invoke(new ForegroundChange(hwnd, ProcessIdOf(hwnd), ClassNameOf(hwnd)));
        }
        catch (Exception ex)
        {
            Core.Log.Error("Foreground handler failed.", ex);
        }
    }

    public void Dispose() => Stop();
}
