using System.Text;

namespace KioskPresenter.Core;

/// <summary>
/// Minimal daily-rolling text log in ProgramData\KioskPresenter\Logs. One line per event,
/// readable in Notepad. Keeps 14 days. Never throws: logging must not take the kiosk down.
/// </summary>
public static class Log
{
    private static readonly object Gate = new();
    private static string? _folder;
    private static string _source = "app";
    private const int KeepDays = 14;

    public static void Init(string folder, string source)
    {
        _folder = folder;
        _source = source;
        try
        {
            Directory.CreateDirectory(folder);
            foreach (var old in Directory.EnumerateFiles(folder, "kiosk-*.log"))
            {
                if (File.GetLastWriteTimeUtc(old) < DateTime.UtcNow.AddDays(-KeepDays)) File.Delete(old);
            }
        }
        catch
        {
            // Best effort only.
        }
    }

    public static void Info(string message) => Write("INFO", message);

    public static void Warn(string message) => Write("WARN", message);

    public static void Error(string message, Exception? ex = null) =>
        Write("ERROR", ex is null ? message : $"{message} | {ex.GetType().Name}: {ex.Message}");

    private static void Write(string level, string message)
    {
        var line = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} [{level}] [{_source}] {message}{Environment.NewLine}";
        System.Diagnostics.Debug.Write(line);
        if (_folder is null) return;
        try
        {
            lock (Gate)
            {
                var file = Path.Combine(_folder, $"kiosk-{DateTime.Now:yyyyMMdd}.log");
                File.AppendAllText(file, line, Encoding.UTF8);
            }
        }
        catch
        {
            // Disk full or folder gone: drop the line rather than crash.
        }
    }

    /// <summary>Most recent lines at ERROR level from today's log, newest last.</summary>
    public static IReadOnlyList<string> RecentErrors(int max = 5)
    {
        if (_folder is null) return [];
        try
        {
            var file = Path.Combine(_folder, $"kiosk-{DateTime.Now:yyyyMMdd}.log");
            if (!File.Exists(file)) return [];
            using var stream = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
            using var reader = new StreamReader(stream);
            var errors = new List<string>();
            while (reader.ReadLine() is { } line)
            {
                if (line.Contains("[ERROR]", StringComparison.Ordinal)) errors.Add(line);
            }
            return errors.TakeLast(max).ToList();
        }
        catch
        {
            return [];
        }
    }
}
