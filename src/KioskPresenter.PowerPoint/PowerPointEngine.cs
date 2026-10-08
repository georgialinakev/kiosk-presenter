using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Threading;
using KioskPresenter.Core;

namespace KioskPresenter.PowerPoint;

/// <summary>What the host learns when a show starts.</summary>
public sealed record ShowStart(IntPtr Hwnd, int ProcessId, int SlideCount);

public enum PollKind
{
    /// <summary>Show is running on <see cref="ShowPoll.SlideIndex"/>.</summary>
    Running,

    /// <summary>Past the last slide (PowerPoint's black "end of slide show" screen).</summary>
    Done,

    /// <summary>No show window any more: Esc in PowerPoint, a crash, or our own teardown.</summary>
    Gone,

    /// <summary>PowerPoint is busy (animation, video loading). Try again next tick.</summary>
    Busy,
}

public readonly record struct ShowPoll(PollKind Kind, int SlideIndex = 0)
{
    public static readonly ShowPoll Done = new(PollKind.Done);
    public static readonly ShowPoll Gone = new(PollKind.Gone);
    public static readonly ShowPoll Busy = new(PollKind.Busy);
    public static ShowPoll At(int slideIndex) => new(PollKind.Running, slideIndex);
}

/// <summary>
/// Drives desktop PowerPoint over late-bound COM. Every call runs on one dedicated STA thread,
/// so a slow or hung PowerPoint never blocks the UI thread. One PowerPoint instance per show;
/// it is quit and released after every show.
/// </summary>
/// <remarks>
/// If a call hangs (a modal PowerPoint dialog, a broken file), the host times out, calls
/// <see cref="Abandon"/> to kill the process, disposes this engine and creates a new one.
/// </remarks>
public sealed class PowerPointEngine : IDisposable
{
    private const int MsoTrue = -1;
    private const int MsoFalse = 0;
    private const int PpAlertsNone = 1;
    private const int PpShowTypeSpeaker = 1;
    private const int PpShowAll = 1;
    private const int PpSlideShowUseSlideTimings = 2;
    private const int PpSlideShowDone = 5;

    private readonly Thread _thread;
    private readonly Dispatcher _dispatcher;
    private dynamic? _app;
    private dynamic? _presentation;
    private dynamic? _window;
    private volatile int _processId;
    private bool _disposed;

    public PowerPointEngine()
    {
        using var ready = new ManualResetEventSlim();
        Dispatcher? dispatcher = null;
        _thread = new Thread(() =>
        {
            dispatcher = Dispatcher.CurrentDispatcher;
            ready.Set();
            Dispatcher.Run();
        })
        {
            IsBackground = true,
            Name = "PowerPoint STA",
        };
        _thread.SetApartmentState(ApartmentState.STA);
        _thread.Start();
        ready.Wait();
        _dispatcher = dispatcher!;
    }

    /// <summary>PowerPoint process of the current show, 0 when none.</summary>
    public int ProcessId => _processId;

    public bool HasShow => _processId != 0;

    /// <summary>Opens the deck read-only with no document window and starts the slideshow.</summary>
    public Task<ShowStart> StartShowAsync(string path) => Run(() =>
    {
        TryRemoveMarkOfTheWeb(path);
        var before = SessionPowerPointIds();
        var type = Type.GetTypeFromProgID(PowerPointInfo.ProgId)
            ?? throw new InvalidOperationException("PowerPoint is not installed.");
        _app = Activator.CreateInstance(type) ?? throw new InvalidOperationException("PowerPoint did not start.");
        _processId = SessionPowerPointIds().Except(before).FirstOrDefault();

        try { _app.DisplayAlerts = PpAlertsNone; } catch (Exception) { /* older builds */ }

        // Open(FileName, ReadOnly, Untitled, WithWindow)
        _presentation = _app.Presentations.Open(path, MsoTrue, MsoFalse, MsoFalse);
        int slideCount = _presentation.Slides.Count;

        dynamic settings = _presentation.SlideShowSettings;
        settings.ShowType = PpShowTypeSpeaker;
        settings.RangeType = PpShowAll;
        settings.LoopUntilStopped = MsoFalse;
        settings.ShowWithAnimation = MsoTrue;
        settings.ShowWithNarration = MsoTrue;
        settings.AdvanceMode = PpSlideShowUseSlideTimings;
        try { settings.ShowPresenterView = MsoFalse; } catch (Exception) { /* not on every build */ }

        _window = settings.Run();
        int hwnd = _window.HWND;
        if (_processId == 0) _processId = Native.ProcessIdOf(new IntPtr(hwnd));
        try { _window.Activate(); } catch (Exception) { /* host foregrounds it anyway */ }
        return new ShowStart(new IntPtr(hwnd), _processId, slideCount);
    });

    public Task<ShowPoll> PollAsync() => Run(() =>
    {
        if (_app is null) return ShowPoll.Gone;
        try
        {
            int windows = _app.SlideShowWindows.Count;
            if (windows == 0) return ShowPoll.Gone;
            dynamic view = _app.SlideShowWindows.Item(1).View;
            int state = view.State;
            if (state == PpSlideShowDone) return ShowPoll.Done;
            int index = view.Slide.SlideIndex;
            return ShowPoll.At(index);
        }
        catch (COMException ex) when (IsDisconnected(ex))
        {
            return ShowPoll.Gone;
        }
        catch (Exception)
        {
            // RPC_E_CALL_REJECTED while PowerPoint is mid-animation, or a transition frame
            // with no current slide. The host gives up only after many busy ticks in a row.
            return ShowPoll.Busy;
        }
    });

    public Task NextAsync() => Run(() => WithView(v => v.Next()));

    public Task PreviousAsync() => Run(() => WithView(v => v.Previous()));

    /// <summary>
    /// Exits the show, closes the deck and quits PowerPoint. Safe to call when nothing is open.
    /// Read <see cref="ProcessId"/> first if you want to make sure the process exits.
    /// </summary>
    public Task TeardownAsync() => Run(() =>
    {
        try
        {
            if (_app is not null)
            {
                int windows = _app.SlideShowWindows.Count;
                if (windows > 0) _app.SlideShowWindows.Item(1).View.Exit();
            }
        }
        catch (Exception) { /* already gone */ }
        try { _presentation?.Close(); } catch (Exception) { /* already closed */ }
        try { _app?.Quit(); } catch (Exception) { /* already quit */ }
        Release(ref _window);
        Release(ref _presentation);
        Release(ref _app);
        GC.Collect();
        GC.WaitForPendingFinalizers();
        _processId = 0;
    });

    /// <summary>Exports slide 1 as a PNG for the library card. Only call when no show is running.</summary>
    public Task ExportThumbnailAsync(string path, string destination, int width = 960, int height = 540) => Run(() =>
    {
        if (_app is not null) throw new InvalidOperationException("A show is running.");
        var before = SessionPowerPointIds();
        dynamic? app = null;
        dynamic? presentation = null;
        try
        {
            TryRemoveMarkOfTheWeb(path);
            var type = Type.GetTypeFromProgID(PowerPointInfo.ProgId)
                ?? throw new InvalidOperationException("PowerPoint is not installed.");
            app = Activator.CreateInstance(type);
            _processId = SessionPowerPointIds().Except(before).FirstOrDefault();
            try { app!.DisplayAlerts = PpAlertsNone; } catch (Exception) { /* older builds */ }
            presentation = app!.Presentations.Open(path, MsoTrue, MsoFalse, MsoFalse);
            Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(destination))!);
            presentation.Slides.Item(1).Export(destination, "PNG", width, height);
        }
        finally
        {
            try { presentation?.Close(); } catch (Exception) { /* ignore */ }
            try { app?.Quit(); } catch (Exception) { /* ignore */ }
            Release(ref presentation);
            Release(ref app);
            GC.Collect();
            GC.WaitForPendingFinalizers();
            _processId = 0;
        }
    });

    /// <summary>
    /// Last resort when a call hangs: kills the show's PowerPoint process so the STA thread's
    /// pending COM call fails and returns. Callable from any thread.
    /// </summary>
    public void Abandon()
    {
        var pid = _processId;
        _processId = 0;
        if (pid != 0) KillProcess(pid);
    }

    /// <summary>Called after teardown: kills the process if PowerPoint didn't exit on its own.</summary>
    public static async Task EnsureExitedAsync(int processId, TimeSpan grace)
    {
        if (processId == 0) return;
        try
        {
            using var process = Process.GetProcessById(processId);
            using var cts = new CancellationTokenSource(grace);
            try
            {
                await process.WaitForExitAsync(cts.Token).ConfigureAwait(false);
            }
            catch (OperationCanceledException)
            {
                Log.Warn($"PowerPoint {processId} still running {grace.TotalSeconds:0} s after quit; killing it.");
                KillProcess(processId);
            }
        }
        catch (ArgumentException)
        {
            // Already gone.
        }
    }

    /// <summary>Kills every PowerPoint in this Windows session. Used once at kiosk start to clear orphans.</summary>
    public static int KillSessionPowerPoints()
    {
        var killed = 0;
        foreach (var id in SessionPowerPointIds())
        {
            if (KillProcess(id)) killed++;
        }
        return killed;
    }

    private Task Run(Action action)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        return _dispatcher.InvokeAsync(action).Task;
    }

    private Task<T> Run<T>(Func<T> func)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        return _dispatcher.InvokeAsync(func).Task;
    }

    private void WithView(Action<dynamic> action)
    {
        if (_app is null) return;
        try
        {
            int windows = _app.SlideShowWindows.Count;
            if (windows == 0) return;
            action(_app.SlideShowWindows.Item(1).View);
        }
        catch (Exception ex)
        {
            Log.Warn($"PowerPoint navigation failed: {ex.Message}");
        }
    }

    private static void Release(ref dynamic? com)
    {
        object? obj = com;
        com = null;
        if (obj is not null && Marshal.IsComObject(obj))
        {
            try { Marshal.FinalReleaseComObject(obj); } catch (Exception) { /* ignore */ }
        }
    }

    private static bool IsDisconnected(COMException ex) =>
        (uint)ex.HResult is 0x800706BA /* RPC server unavailable */ or 0x80010108 /* object disconnected */ or 0x800706BE /* RPC call failed */;

    private static IEnumerable<int> SessionPowerPointIds()
    {
        var session = Process.GetCurrentProcess().SessionId;
        var ids = new List<int>();
        foreach (var p in Process.GetProcessesByName("POWERPNT"))
        {
            using (p)
            {
                try
                {
                    if (p.SessionId == session) ids.Add(p.Id);
                }
                catch (Exception) { /* exited meanwhile */ }
            }
        }
        return ids;
    }

    private static bool KillProcess(int processId)
    {
        try
        {
            using var process = Process.GetProcessById(processId);
            process.Kill(entireProcessTree: true);
            process.WaitForExit(3000);
            return true;
        }
        catch (Exception)
        {
            return false;
        }
    }

    /// <summary>
    /// Files that came by email or download carry a Mark of the Web, which makes PowerPoint open
    /// them in Protected View. Removing the Zone.Identifier stream is what "Unblock" does.
    /// </summary>
    private static void TryRemoveMarkOfTheWeb(string path)
    {
        try
        {
            Native.DeleteFile(path + ":Zone.Identifier");
        }
        catch (Exception)
        {
            // No permission or no stream: PowerPoint will tell us if it matters.
        }
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        try
        {
            _dispatcher.InvokeShutdown();
        }
        catch (Exception) { /* thread already gone */ }
    }
}
