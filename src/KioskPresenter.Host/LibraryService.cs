using System.IO;
using System.Windows.Threading;
using KioskPresenter.Core;

namespace KioskPresenter.Host;

/// <summary>
/// Keeps the library current: scans the content folder on start, on file changes (debounced)
/// and every 30 s as a backstop, and works out each deck's thumbnail.
/// </summary>
internal sealed class LibraryService : IDisposable
{
    public const string CacheHost = "cache.kiosk.local";

    private readonly KioskPaths _paths;
    private readonly Func<string> _contentFolder;
    private readonly DispatcherTimer _debounce;
    private readonly DispatcherTimer _backstop;
    private FileSystemWatcher? _watcher;
    private string _signature = string.Empty;
    private bool _scanning;

    public LibraryService(KioskPaths paths, Func<string> contentFolder)
    {
        _paths = paths;
        _contentFolder = contentFolder;
        _debounce = new DispatcherTimer { Interval = TimeSpan.FromSeconds(2) };
        _debounce.Tick += async (_, _) => { _debounce.Stop(); await RescanAsync(); };
        _backstop = new DispatcherTimer { Interval = TimeSpan.FromSeconds(30) };
        _backstop.Tick += async (_, _) => await RescanAsync();
    }

    public LibraryScan Current { get; private set; } = new(LibraryStatus.Empty, [], [], string.Empty);

    public DateTime LastScan { get; private set; }

    /// <summary>Raised on the UI thread when the deck list or any thumbnail changed.</summary>
    public event Action? Changed;

    public async Task StartAsync()
    {
        await RescanAsync(force: true);
        Watch();
        _backstop.Start();
    }

    public async Task RescanAsync(bool force = false)
    {
        if (_scanning) return;
        _scanning = true;
        try
        {
            var folder = _contentFolder();
            var scan = await Task.Run(() => DeckScanner.Scan(folder));
            await Task.Run(() => PrepareThumbnails(scan));
            LastScan = DateTime.Now;
            foreach (var issue in scan.Issues) Log.Warn($"Skipped {Path.GetFileName(issue.FilePath)}: {issue.Reason}");

            var signature = Signature(scan);
            Current = scan;
            if (force || signature != _signature)
            {
                _signature = signature;
                Log.Info($"Library: {scan.Status}, {scan.Decks.Count} deck(s) in {scan.Folder}");
                Changed?.Invoke();
            }
            if (_watcher is null || !string.Equals(_watcher.Path, scan.Folder, StringComparison.OrdinalIgnoreCase)) Watch();
        }
        catch (Exception ex)
        {
            Log.Error("Library scan failed.", ex);
        }
        finally
        {
            _scanning = false;
        }
    }

    /// <summary>URL the WebView loads for a deck's card image, or null for the design's placeholder.</summary>
    public string? ThumbnailUrl(DeckInfo deck)
    {
        foreach (var ext in new[] { ".png", ".jpg" })
        {
            var file = Path.Combine(_paths.ThumbnailFolder, deck.CacheKey + ext);
            if (File.Exists(file)) return $"https://{CacheHost}/thumbs/{deck.CacheKey}{ext}";
        }
        return null;
    }

    /// <summary>Decks that still only have PowerPoint's small embedded preview, or none.</summary>
    public IEnumerable<DeckInfo> DecksNeedingExport() =>
        Current.Decks.Where(d => !File.Exists(ExportPath(d)) && !FailedExports.Contains(d.CacheKey));

    public string ExportPath(DeckInfo deck) => Path.Combine(_paths.ThumbnailFolder, deck.CacheKey + ".png");

    /// <summary>Exports that failed once are not retried until the file changes.</summary>
    public HashSet<string> FailedExports { get; } = new(StringComparer.Ordinal);

    /// <summary>Call after a COM export lands so the UI picks up the sharper image.</summary>
    public void ThumbnailUpdated()
    {
        _signature = Signature(Current);
        Changed?.Invoke();
    }

    private void PrepareThumbnails(LibraryScan scan)
    {
        Directory.CreateDirectory(_paths.ThumbnailFolder);
        foreach (var deck in scan.Decks)
        {
            var png = Path.Combine(_paths.ThumbnailFolder, deck.CacheKey + ".png");
            var jpg = Path.Combine(_paths.ThumbnailFolder, deck.CacheKey + ".jpg");
            if (File.Exists(png) || File.Exists(jpg)) continue;
            if (deck.Format == DeckFormat.OpenXml) DeckReader.TryExtractThumbnail(deck.FilePath, jpg);
        }
    }

    private string Signature(LibraryScan scan) =>
        scan.Status + "|" + string.Join("|", scan.Decks.Select(d => d.CacheKey + ":" + (ThumbnailUrl(d) ?? "-")));

    private void Watch()
    {
        _watcher?.Dispose();
        _watcher = null;
        var folder = _contentFolder();
        if (!Directory.Exists(folder)) return;
        try
        {
            _watcher = new FileSystemWatcher(folder)
            {
                IncludeSubdirectories = false,
                NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite | NotifyFilters.Size,
            };
            _watcher.Created += OnFileEvent;
            _watcher.Changed += OnFileEvent;
            _watcher.Deleted += OnFileEvent;
            _watcher.Renamed += OnFileEvent;
            _watcher.Error += (_, e) => Log.Warn($"Folder watcher error: {e.GetException().Message}");
            _watcher.EnableRaisingEvents = true;
        }
        catch (Exception ex)
        {
            Log.Warn($"Can't watch {folder}: {ex.Message}. The 30 s rescan still runs.");
        }
    }

    private void OnFileEvent(object sender, FileSystemEventArgs e)
    {
        // Watcher events arrive on a thread-pool thread; restart the debounce on the UI thread.
        _debounce.Dispatcher.InvokeAsync(() =>
        {
            _debounce.Stop();
            _debounce.Start();
        });
    }

    public void Dispose()
    {
        _debounce.Stop();
        _backstop.Stop();
        _watcher?.Dispose();
    }
}
