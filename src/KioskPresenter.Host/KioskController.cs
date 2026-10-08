using System.Diagnostics;
using System.IO;
using System.Net.NetworkInformation;
using System.Security.Principal;
using System.Windows;
using System.Windows.Threading;
using KioskPresenter.Core;
using KioskPresenter.Host.Interop;
using KioskPresenter.PowerPoint;
using Microsoft.Win32;

namespace KioskPresenter.Host;

/// <summary>
/// The host's brain. Turns UI commands into PowerPoint actions, keeps the slideshow, the control
/// strip and the UI window in the right order, and keeps visitors inside the kiosk.
/// </summary>
internal sealed class KioskController : IDisposable
{
    private static readonly TimeSpan OpenTimeout = TimeSpan.FromSeconds(60);
    private static readonly TimeSpan CallTimeout = TimeSpan.FromSeconds(10);
    private static readonly TimeSpan ExportTimeout = TimeSpan.FromSeconds(45);
    private static readonly TimeSpan ActivityThrottle = TimeSpan.FromSeconds(1);
    private const int MaxBusyTicks = 40; // 40 x 250 ms = 10 s of PowerPoint not answering
    private static readonly string[] BrowserProcesses = ["msedge", "chrome", "firefox", "iexplore", "brave", "opera"];

    private readonly HostOptions _options;
    private readonly KioskPaths _paths;
    private readonly MainWindow _main;
    private readonly ControlStrip _strip = new();
    private readonly KeyboardHook _keyboard;
    private readonly ForegroundWatcher _foreground = new();
    private readonly InputMonitor _input = new(TimeSpan.FromMilliseconds(250));
    private readonly DispatcherTimer _poll = new() { Interval = TimeSpan.FromMilliseconds(250) };
    private readonly DispatcherTimer _watchdogCheck = new() { Interval = TimeSpan.FromSeconds(30) };
    private readonly DispatcherTimer _refocus = new() { Interval = TimeSpan.FromMilliseconds(150) };
    private readonly DateTime _startedAt = DateTime.Now;
    private readonly Queue<DateTime> _refocusLog = new();

    private KioskConfig _config = new();
    private string? _configError;
    private LibraryService? _library;
    private PowerPointEngine _engine = new();
    private Session? _session;
    private bool _uiReady;
    private bool _openInFlight;
    private bool _cancelOpen;
    private bool _promptVisible;
    private bool _polling;
    private bool _exporting;
    private bool _exiting;
    private int _busyTicks;
    private long _stripHideAt;
    private DateTime _lastActivityPost = DateTime.MinValue;
    private Action? _pendingRefocus;
    private DateTime _refocusPausedUntil = DateTime.MinValue;

    public KioskController(HostOptions options, KioskPaths paths, MainWindow main)
    {
        _options = options;
        _paths = paths;
        _main = main;
        _keyboard = new KeyboardHook(() => _session is not null);
        _poll.Tick += async (_, _) => await PollAsync();
        _watchdogCheck.Tick += (_, _) => Watchdog.EnsureRunning(_options);
        _refocus.Tick += (_, _) =>
        {
            _refocus.Stop();
            var action = _pendingRefocus;
            _pendingRefocus = null;
            action?.Invoke();
        };
    }

    private sealed class Session(DeckInfo deck, ShowStart start, ShowMap map)
    {
        public DeckInfo Deck { get; } = deck;
        public IntPtr Hwnd { get; } = start.Hwnd;
        public int ProcessId { get; } = start.ProcessId;
        public ShowMap Map { get; } = map;
        public int Total { get; } = map.VisibleTotal(start.SlideCount);
        public int LastIndex { get; set; } = -1;
        public bool OnVideo { get; set; }
    }

    // ------------------------------------------------------------------ start / stop

    public async Task StartAsync()
    {
        try
        {
            _paths.EnsureCreated();
        }
        catch (Exception ex)
        {
            Log.Error($"Can't create {_paths.DataRoot}. Was Setup run?", ex);
        }
        TryDelete(_paths.AdminExitFlag);
        LoadConfig();

        var uiFolder = Path.Combine(AppContext.BaseDirectory, "ui");
        if (!File.Exists(Path.Combine(uiFolder, "index.html")))
        {
            throw new FileNotFoundException($"UI files missing: {uiFolder}\\index.html");
        }

        if (_options.KioskMode)
        {
            var killed = PowerPointEngine.KillSessionPowerPoints();
            if (killed > 0) Log.Warn($"Closed {killed} PowerPoint process(es) left over from before.");
        }

        _main.CommandReceived += OnCommand;
        _strip.HomeRequested += () => _ = EndSessionAsync(postHome: true);
        _strip.PreviousRequested += () => _ = SafeAsync(() => _engine.PreviousAsync(), "previous");
        _strip.NextRequested += () => _ = SafeAsync(() => _engine.NextAsync(), "next");
        _keyboard.AdminChord += OnAdminChord;
        _foreground.Changed += OnForeground;
        _input.Input += OnInput;
        _input.Start();

        if (_options.KioskMode)
        {
            _keyboard.Install();
            _foreground.Start();
            WindowHelper.KeepDisplayOn(true);
        }
        if (_options.UseWatchdog)
        {
            Watchdog.EnsureRunning(_options);
            _watchdogCheck.Start();
        }

        _library = new LibraryService(_paths, () => _config.ResolveContentFolder(_paths));
        _library.Changed += () =>
        {
            if (_uiReady) PostLibrary();
            QueueThumbnails();
        };

        await _main.InitializeWebViewAsync(_paths, uiFolder, ConfigScript(), _options.DevTools);
        await _library.StartAsync();
        Log.Info("Kiosk ready.");
    }

    /// <summary>Sign-out or shutdown: tell the watchdog not to restart us.</summary>
    public void OnSessionEnding()
    {
        Log.Info("Windows session ending.");
        WriteExitFlag();
    }

    public void Dispose()
    {
        _poll.Stop();
        _watchdogCheck.Stop();
        _keyboard.Dispose();
        _foreground.Dispose();
        _input.Dispose();
        _library?.Dispose();
        WindowHelper.KeepDisplayOn(false);
        try
        {
            _engine.TeardownAsync().Wait(TimeSpan.FromSeconds(5));
        }
        catch (Exception)
        {
            _engine.Abandon();
        }
        _engine.Dispose();
    }

    // ------------------------------------------------------------------ UI commands

    private void OnCommand(UiCommand command)
    {
        switch (command.Cmd)
        {
            case "ready":
                _uiReady = true;
                PostLibrary();
                if (_session is not null) _ = EndSessionAsync(postHome: true); // UI reloaded mid-show
                break;
            case "open":
            case "replay":
                _ = OpenAsync(command.Id);
                break;
            case "home":
                if (_openInFlight) _cancelOpen = true;
                _ = EndSessionAsync(postHome: false);
                break;
            case "next":
                _ = SafeAsync(() => _engine.NextAsync(), "next");
                break;
            case "prev":
                _ = SafeAsync(() => _engine.PreviousAsync(), "previous");
                break;
            case "prompt":
                SetPrompt(command.Visible == true);
                break;
            case "pinSubmit":
                PinSubmit(command.Digits);
                break;
            case "setPin":
                _ = SetPinAsync(command.Digits);
                break;
            case "admin" when command.Action == "reload":
                _ = ReloadAsync();
                break;
            case "admin" when command.Action == "exit":
                ExitKiosk();
                break;
            default:
                Log.Warn($"Unknown UI command: {command.Cmd} {command.Action}");
                break;
        }
    }

    // ------------------------------------------------------------------ shows

    private async Task OpenAsync(string? id)
    {
        if (_session is not null || _openInFlight || _library is null) return;
        var deck = _library.Current.Find(id);
        if (deck is null)
        {
            _main.Post(Bridge.OpenFailed(id ?? string.Empty, "Presentation not found."));
            return;
        }

        _openInFlight = true;
        _cancelOpen = false;
        _main.Post(Bridge.Opening(deck.Id));
        Log.Info($"Opening {Path.GetFileName(deck.FilePath)}");
        try
        {
            await WaitForExportAsync();
            var start = await _engine.StartShowAsync(deck.FilePath).WaitAsync(OpenTimeout);
            if (_cancelOpen)
            {
                Log.Info("Visitor went back while the deck was opening.");
                await TeardownEngineAsync();
                return;
            }

            var session = new Session(deck, start, deck.ToShowMap());
            _session = session;
            _busyTicks = 0;
            _promptVisible = false;

            WindowHelper.SetTopmost(_main.Handle, false);
            WindowHelper.ForceForeground(start.Hwnd);
            _strip.SetCounter(1, session.Total);
            _strip.ShowStrip();
            _stripHideAt = Environment.TickCount64 + (long)(_config.Timers.StripHide * 1000);

            _main.Post(Bridge.ShowStarted(deck.Id, session.Total, session.Map.IsVideo(1)));
            _poll.Start();
            Log.Info($"Show started: {session.Total} slide(s), PowerPoint pid {start.ProcessId}.");
        }
        catch (TimeoutException)
        {
            Log.Error($"PowerPoint took longer than {OpenTimeout.TotalSeconds:0} s to open {deck.FilePath}; abandoning it.");
            RecycleEngine();
            FailOpen(deck, "Timed out");
        }
        catch (Exception ex)
        {
            Log.Error($"Could not open {deck.FilePath}.", ex);
            await TeardownEngineAsync();
            FailOpen(deck, ex.Message);
        }
        finally
        {
            _openInFlight = false;
        }
    }

    private void FailOpen(DeckInfo deck, string reason)
    {
        _main.Post(Bridge.OpenFailed(deck.Id, reason));
        BringMainToFront();
        QueueThumbnails();
    }

    private async Task PollAsync()
    {
        if (_polling) return;
        var session = _session;
        if (session is null)
        {
            _poll.Stop();
            return;
        }

        _polling = true;
        try
        {
            ShowPoll poll;
            try
            {
                poll = await _engine.PollAsync().WaitAsync(TimeSpan.FromSeconds(5));
            }
            catch (TimeoutException)
            {
                poll = ShowPoll.Busy;
            }

            if (!ReferenceEquals(session, _session)) return; // ended while we waited

            switch (poll.Kind)
            {
                case PollKind.Busy:
                    if (++_busyTicks > MaxBusyTicks)
                    {
                        Log.Warn("PowerPoint stopped answering during the show; returning to the library.");
                        await EndSessionAsync(postHome: true);
                    }
                    break;
                case PollKind.Gone:
                    Log.Info("Slideshow closed by PowerPoint (Esc, crash or a link).");
                    await EndSessionAsync(postHome: true);
                    break;
                case PollKind.Done:
                    Log.Info("Reached the end of the presentation.");
                    await EndSessionAsync(postHome: false, showEnded: true);
                    break;
                case PollKind.Running:
                    _busyTicks = 0;
                    if (poll.SlideIndex != session.LastIndex)
                    {
                        session.LastIndex = poll.SlideIndex;
                        var position = session.Map.Position(poll.SlideIndex);
                        session.OnVideo = session.Map.IsVideo(poll.SlideIndex);
                        _strip.SetCounter(position, session.Total);
                        _main.Post(Bridge.Slide(position, session.Total, session.OnVideo));
                    }
                    break;
            }

            if (_session is null) return;

            if (_strip.IsShowing && !_promptVisible && Environment.TickCount64 > _stripHideAt) _strip.HideStrip();

            // A looping or very long video can't hold the kiosk forever.
            if (session.OnVideo && Environment.TickCount64 - _input.LastInputMs > _config.Timers.VideoIdleCap * 1000)
            {
                Log.Info("Video slide idle past the cap; returning to the library.");
                await EndSessionAsync(postHome: true);
            }
        }
        catch (Exception ex)
        {
            Log.Error("Show poll failed.", ex);
        }
        finally
        {
            _polling = false;
        }
    }

    private async Task EndSessionAsync(bool postHome, bool showEnded = false)
    {
        var session = _session;
        if (session is null) return;
        _session = null;
        _poll.Stop();
        _promptVisible = false;
        _strip.Retire();

        BringMainToFront();
        if (showEnded) _main.Post(Bridge.ShowEnded());
        else if (postHome) _main.Post(Bridge.Home());

        await TeardownEngineAsync();
        Log.Info($"Show closed: {Path.GetFileName(session.Deck.FilePath)}");
        QueueThumbnails();
    }

    private async Task TeardownEngineAsync()
    {
        var pid = _engine.ProcessId;
        try
        {
            await _engine.TeardownAsync().WaitAsync(CallTimeout);
        }
        catch (TimeoutException)
        {
            Log.Warn("PowerPoint didn't close in time; killing it.");
            RecycleEngine();
        }
        catch (Exception ex)
        {
            Log.Warn($"PowerPoint teardown: {ex.Message}");
        }
        await PowerPointEngine.EnsureExitedAsync(pid, TimeSpan.FromSeconds(5));
    }

    /// <summary>Kills a wedged PowerPoint and starts a fresh engine thread.</summary>
    private void RecycleEngine()
    {
        var old = _engine;
        old.Abandon();
        _engine = new PowerPointEngine();
        old.Dispose();
    }

    private void SetPrompt(bool visible)
    {
        _promptVisible = visible;
        var session = _session;
        if (session is null) return;
        if (visible)
        {
            // Still there? lives in the UI window, which sits behind the slideshow during a show.
            _strip.HideStrip();
            WindowHelper.SetTopmost(_main.Handle, _options.KioskMode);
            WindowHelper.ForceForeground(_main.Handle);
            _main.FocusUi();
        }
        else
        {
            WindowHelper.SetTopmost(_main.Handle, false);
            WindowHelper.ForceForeground(session.Hwnd);
            ShowStripFor(_config.Timers.StripHide);
        }
    }

    private void ShowStripFor(double seconds)
    {
        _strip.ShowStrip();
        _stripHideAt = Environment.TickCount64 + (long)(seconds * 1000);
    }

    private void OnInput()
    {
        if (_session is null || _promptVisible) return;
        ShowStripFor(_config.Timers.StripHide);
        if (DateTime.UtcNow - _lastActivityPost >= ActivityThrottle)
        {
            _lastActivityPost = DateTime.UtcNow;
            _main.Post(Bridge.Activity());
        }
    }

    private void OnAdminChord()
    {
        _ = OpenPinAsync();

        async Task OpenPinAsync()
        {
            if (_session is not null) await EndSessionAsync(postHome: true);
            BringMainToFront();
            _main.Post(Bridge.Pin());
        }
    }

    // ------------------------------------------------------------------ keeping visitors in

    private void OnForeground(ForegroundChange change)
    {
        if (!_options.KioskMode || _exiting) return;
        if (change.ProcessId == Environment.ProcessId) return;

        var session = _session;
        if (session is not null && !_promptVisible)
        {
            if (change.Hwnd == session.Hwnd) return;
            if (change.ProcessId == session.ProcessId && change.ClassName == "screenClass") return;
            CloseStrayBrowser(change);
            ScheduleRefocus(() =>
            {
                if (_session is not { } s) return;
                if (NativeMethods.IsWindow(s.Hwnd)) WindowHelper.ForceForeground(s.Hwnd);
                else _ = EndSessionAsync(postHome: true);
            });
            return;
        }

        CloseStrayBrowser(change);
        ScheduleRefocus(BringMainToFront);
    }

    private void ScheduleRefocus(Action action)
    {
        var now = DateTime.UtcNow;
        if (now < _refocusPausedUntil) return;
        _refocusLog.Enqueue(now);
        while (_refocusLog.Count > 0 && now - _refocusLog.Peek() > TimeSpan.FromSeconds(5)) _refocusLog.Dequeue();
        if (_refocusLog.Count > 12)
        {
            // Something else keeps grabbing the screen (a system prompt, say). Don't fight it in a loop.
            Log.Warn("Focus is being taken repeatedly; pausing the focus guard for 5 s.");
            _refocusPausedUntil = now.AddSeconds(5);
            _refocusLog.Clear();
            return;
        }
        _pendingRefocus = action;
        _refocus.Stop();
        _refocus.Start();
    }

    private void CloseStrayBrowser(ForegroundChange change)
    {
        try
        {
            using var process = Process.GetProcessById(change.ProcessId);
            if (!BrowserProcesses.Contains(process.ProcessName, StringComparer.OrdinalIgnoreCase)) return;
            Log.Warn($"Closing a {process.ProcessName} window opened during the kiosk session.");
            NativeMethods.PostMessage(change.Hwnd, NativeMethods.WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
        }
        catch (Exception)
        {
            // Process gone already.
        }
    }

    private void BringMainToFront()
    {
        WindowHelper.SetTopmost(_main.Handle, _options.KioskMode);
        WindowHelper.ForceForeground(_main.Handle);
        _main.FocusUi();
    }

    // ------------------------------------------------------------------ admin

    private void PinSubmit(string? digits)
    {
        if (PinHasher.Verify(digits, _config.PinHash))
        {
            Log.Info("Admin PIN accepted.");
            _main.Post(Bridge.AdminStatus(BuildStatus()));
            _main.Post(Bridge.PinResult(true));
        }
        else
        {
            Log.Warn("Wrong admin PIN entered.");
            _main.Post(Bridge.PinResult(false));
        }
    }

    private async Task SetPinAsync(string? digits)
    {
        if (!_config.NeedsSetup)
        {
            Log.Warn("Ignored setPin: a PIN is already set.");
            _main.Post(Bridge.SetupResult(false));
            return;
        }
        if (digits is null || !PinHasher.IsValidPin(digits))
        {
            _main.Post(Bridge.SetupResult(false));
            return;
        }
        try
        {
            _config.PinHash = PinHasher.Hash(digits);
            ConfigStore.Save(_paths.ConfigFile, _config);
            await _main.UpdateConfigScriptAsync(ConfigScript());
            Log.Info("Admin PIN set.");
            _main.Post(Bridge.SetupResult(true));
        }
        catch (Exception ex)
        {
            Log.Error("Could not save the admin PIN.", ex);
            _main.Post(Bridge.SetupResult(false));
        }
    }

    private async Task ReloadAsync()
    {
        Log.Info("Admin: reload library.");
        LoadConfig();
        await _main.UpdateConfigScriptAsync(ConfigScript());
        if (_library is not null)
        {
            _library.FailedExports.Clear();
            await _library.RescanAsync(force: true);
        }
        _main.Post(Bridge.AdminStatus(BuildStatus()));
    }

    private void ExitKiosk()
    {
        Log.Info("Admin: exit kiosk.");
        _exiting = true;
        WriteExitFlag();
        _keyboard.Uninstall();
        _foreground.Stop();
        _ = Shutdown();

        async Task Shutdown()
        {
            await EndSessionAsync(postHome: false);
            WindowHelper.KeepDisplayOn(false);
            _main.AllowClose = true;
            Application.Current.Shutdown(0);
        }
    }

    private AdminStatusDto BuildStatus()
    {
        var scan = _library?.Current;
        var folder = _config.ResolveContentFolder(_paths);
        var folderOk = Directory.Exists(folder);
        var ppt = PowerPointInfo.Describe();
        var network = NetworkInterface.GetIsNetworkAvailable();
        var warnings = new List<AdminWarning>();

        if (_configError is not null) warnings.Add(new AdminWarning("config", "Settings file problem", _configError));
        if (ppt is null) warnings.Add(new AdminWarning("ppt", "PowerPoint not found", "Install Microsoft PowerPoint for this Windows user, then reload."));
        if (!folderOk) warnings.Add(new AdminWarning("folder", "Content folder missing", $"{folder} does not exist or can't be read by {Account()}."));
        if (scan is { Issues.Count: > 0 })
        {
            var names = string.Join("; ", scan.Issues.Take(3).Select(i => $"{Path.GetFileName(i.FilePath)}: {i.Reason}"));
            var more = scan.Issues.Count > 3 ? $" (+{scan.Issues.Count - 3} more)" : string.Empty;
            warnings.Add(new AdminWarning("skipped", $"{scan.Issues.Count} file(s) skipped", names + more));
        }
        if (!network && scan is not null && scan.Decks.Any(d => d.HasVideo))
        {
            warnings.Add(new AdminWarning("net", "Offline", "Decks with online video (YouTube) won't play their video until the network is back."));
        }

        return new AdminStatusDto(
            Version: AppVersion.Current,
            Folder: folder,
            FolderOk: folderOk,
            Decks: scan?.Decks.Count ?? 0,
            Ppt: ppt,
            Network: network,
            LastRefresh: _library?.LastScan.ToString("HH:mm:ss") ?? "-",
            Account: Account(),
            Autostart: AutoStartIsThisUser(),
            Warnings: warnings);
    }

    private static string Account() => $"{Environment.UserDomainName}\\{Environment.UserName}";

    /// <summary>Setup records which user the auto-start task is for.</summary>
    private static bool AutoStartIsThisUser()
    {
        try
        {
            using var key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\KioskPresenter");
            var sid = key?.GetValue("AutoStartSid") as string;
            return sid is not null && string.Equals(sid, WindowsIdentity.GetCurrent().User?.Value, StringComparison.OrdinalIgnoreCase);
        }
        catch (Exception)
        {
            return false;
        }
    }

    // ------------------------------------------------------------------ library + thumbnails

    private void PostLibrary()
    {
        if (_library is null) return;
        _main.Post(Bridge.Library(_library.Current, _library.ThumbnailUrl));
    }

    private Task? _exportTask;

    private async Task WaitForExportAsync()
    {
        var task = _exportTask;
        if (task is null || task.IsCompleted) return;
        try
        {
            await task.WaitAsync(ExportTimeout);
        }
        catch (Exception)
        {
            // The export logs its own failure.
        }
    }

    /// <summary>Exports sharper thumbnails one at a time, only while nobody is watching a show.</summary>
    private void QueueThumbnails()
    {
        if (_exporting || _session is not null || _openInFlight || _exiting || _library is null) return;
        if (!PowerPointInfo.IsInstalled) return;
        var deck = _library.DecksNeedingExport().FirstOrDefault(d => d.Format == DeckFormat.OpenXml || d.Format == DeckFormat.Legacy);
        if (deck is null) return;
        _exporting = true;
        _exportTask = ExportAsync(deck);
    }

    private async Task ExportAsync(DeckInfo deck)
    {
        var library = _library!;
        var final = library.ExportPath(deck);
        var temp = final + ".tmp.png";
        try
        {
            await _engine.ExportThumbnailAsync(deck.FilePath, temp).WaitAsync(ExportTimeout);
            File.Move(temp, final, overwrite: true);
            Log.Info($"Thumbnail ready: {Path.GetFileName(deck.FilePath)}");
            library.ThumbnailUpdated();
        }
        catch (TimeoutException)
        {
            Log.Warn($"Thumbnail export timed out: {Path.GetFileName(deck.FilePath)}");
            library.FailedExports.Add(deck.CacheKey);
            RecycleEngine();
        }
        catch (Exception ex)
        {
            Log.Warn($"Thumbnail export failed for {Path.GetFileName(deck.FilePath)}: {ex.Message}");
            library.FailedExports.Add(deck.CacheKey);
        }
        finally
        {
            TryDelete(temp);
            _exporting = false;
            await Task.Delay(500);
            QueueThumbnails();
        }
    }

    // ------------------------------------------------------------------ helpers

    private void LoadConfig()
    {
        _config = ConfigStore.Load(_paths.ConfigFile, out _configError);
        if (_configError is not null) Log.Error(_configError);
        if (!File.Exists(_paths.ConfigFile))
        {
            try
            {
                ConfigStore.Save(_paths.ConfigFile, _config);
            }
            catch (Exception ex)
            {
                Log.Warn($"Could not write a default config.json: {ex.Message}");
            }
        }
    }

    private string ConfigScript()
    {
        string? logoUrl = null;
        var logo = _config.Brand.Logo;
        if (!string.IsNullOrWhiteSpace(logo) && File.Exists(Path.Combine(_paths.BrandFolder, logo)))
        {
            logoUrl = $"https://{MainWindow.BrandHost}/{Uri.EscapeDataString(logo)}";
        }
        return UiConfigScript.Build(_config, logoUrl);
    }

    private void WriteExitFlag()
    {
        try
        {
            Directory.CreateDirectory(_paths.StateFolder);
            File.WriteAllText(_paths.AdminExitFlag, DateTime.Now.ToString("O"));
        }
        catch (Exception ex)
        {
            Log.Warn($"Could not write the exit flag: {ex.Message}");
        }
    }

    private static void TryDelete(string path)
    {
        try
        {
            if (File.Exists(path)) File.Delete(path);
        }
        catch (Exception)
        {
            // Best effort.
        }
    }

    private static async Task SafeAsync(Func<Task> action, string what)
    {
        try
        {
            await action().WaitAsync(CallTimeout);
        }
        catch (Exception ex)
        {
            Log.Warn($"PowerPoint {what} failed: {ex.Message}");
        }
    }
}
