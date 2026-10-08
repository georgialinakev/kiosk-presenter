namespace KioskPresenter.Core;

/// <summary>
/// Every on-disk location the app uses, derived from one data root so tests can point it anywhere.
/// </summary>
public sealed class KioskPaths
{
    public KioskPaths(string dataRoot)
    {
        if (string.IsNullOrWhiteSpace(dataRoot)) throw new ArgumentException("Data root is required.", nameof(dataRoot));
        DataRoot = dataRoot;
    }

    /// <summary>C:\ProgramData\KioskPresenter on a real install.</summary>
    public static KioskPaths Default { get; } = new(
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "KioskPresenter"));

    public string DataRoot { get; }

    /// <summary>Where presentations live unless config.json says otherwise.</summary>
    public string DefaultContentFolder => Path.Combine(DataRoot, "Content");

    public string ConfigFile => Path.Combine(DataRoot, "config.json");

    /// <summary>Optional logo and other brand files, served to the UI.</summary>
    public string BrandFolder => Path.Combine(DataRoot, "Brand");

    public string CacheFolder => Path.Combine(DataRoot, "Cache");

    public string ThumbnailFolder => Path.Combine(CacheFolder, "thumbs");

    public string LogFolder => Path.Combine(DataRoot, "Logs");

    public string StateFolder => Path.Combine(DataRoot, "State");

    /// <summary>Written by an admin exit so the watchdog stands down instead of relaunching.</summary>
    public string AdminExitFlag => Path.Combine(StateFolder, "admin-exit.flag");

    /// <summary>The WebView2 user data folder (cookies, cache). Per machine, not per user.</summary>
    public string WebViewDataFolder => Path.Combine(CacheFolder, "WebView2");

    public void EnsureCreated()
    {
        foreach (var dir in new[] { DataRoot, DefaultContentFolder, BrandFolder, CacheFolder, ThumbnailFolder, LogFolder, StateFolder })
        {
            Directory.CreateDirectory(dir);
        }
    }
}
