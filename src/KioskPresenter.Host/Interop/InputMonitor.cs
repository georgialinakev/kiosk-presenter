using System.Windows.Threading;

namespace KioskPresenter.Host.Interop;

/// <summary>
/// Notices any keyboard, mouse or touch input in the session, wherever it lands. During a show
/// input goes to PowerPoint, so this is how the strip and the UI's idle timer learn about it.
/// </summary>
internal sealed class InputMonitor : IDisposable
{
    private readonly DispatcherTimer _timer;
    private uint _lastTick;

    public InputMonitor(TimeSpan interval)
    {
        _timer = new DispatcherTimer(DispatcherPriority.Background) { Interval = interval };
        _timer.Tick += (_, _) => Check();
    }

    /// <summary>Raised at most once per interval when there has been input since the last check.</summary>
    public event Action? Input;

    /// <summary>When the last input happened, as Environment.TickCount64-compatible milliseconds.</summary>
    public long LastInputMs { get; private set; } = Environment.TickCount64;

    public void Start()
    {
        _lastTick = NativeMethods.LastInputTick();
        LastInputMs = Environment.TickCount64;
        _timer.Start();
    }

    public void Stop() => _timer.Stop();

    private void Check()
    {
        var tick = NativeMethods.LastInputTick();
        if (tick == 0 || tick == _lastTick) return;
        _lastTick = tick;
        LastInputMs = Environment.TickCount64;
        Input?.Invoke();
    }

    public void Dispose() => _timer.Stop();
}
