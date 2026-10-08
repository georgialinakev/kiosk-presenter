using System.Text.Json;
using System.Text.Json.Serialization;

namespace KioskPresenter.Core;

/// <summary>
/// C:\ProgramData\KioskPresenter\config.json. Everything client-specific lives here, never in the repo.
/// </summary>
public sealed class KioskConfig
{
    /// <summary>Folder scanned for presentations. Null means <see cref="KioskPaths.DefaultContentFolder"/>.</summary>
    public string? ContentFolder { get; set; }

    public BrandConfig Brand { get; set; } = new();

    public TimerConfig Timers { get; set; } = new();

    /// <summary>Fades and scales in the UI. Off for slow hardware.</summary>
    public bool Motion { get; set; } = true;

    /// <summary>PBKDF2 hash from <see cref="PinHasher"/>. Null until first-run setup.</summary>
    public string? PinHash { get; set; }

    [JsonIgnore]
    public bool NeedsSetup => string.IsNullOrEmpty(PinHash);

    public string ResolveContentFolder(KioskPaths paths) =>
        string.IsNullOrWhiteSpace(ContentFolder) ? paths.DefaultContentFolder : Environment.ExpandEnvironmentVariables(ContentFolder);
}

public sealed class BrandConfig
{
    public string Title { get; set; } = "Welcome";

    public string Subtitle { get; set; } = "Choose a presentation to begin";

    /// <summary>File name inside the Brand folder, e.g. "logo.png". Null shows the empty logo slot.</summary>
    public string? Logo { get; set; }

    /// <summary>Hex colour such as "#4fd1b5". Null keeps the design's accent.</summary>
    public string? Accent { get; set; }
}

/// <summary>All values in seconds. Names match the UI's KIOSK_CONFIG.timers.</summary>
public sealed class TimerConfig
{
    public double LibraryIdle { get; set; } = 120;
    public double PlayerIdle { get; set; } = 180;
    public double Prompt { get; set; } = 15;
    public double EndAutoReturn { get; set; } = 15;
    public double OpeningSlow { get; set; } = 10;
    public double StripHide { get; set; } = 4;
    public double CursorHide { get; set; } = 3;
    public double InputLock { get; set; } = 1.5;
    public double AdminHold { get; set; } = 5;

    /// <summary>Host-side cap on the video idle exemption, so a looping video can't hold the kiosk forever.</summary>
    public double VideoIdleCap { get; set; } = 900;
}

public static class ConfigStore
{
    public static readonly JsonSerializerOptions Json = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = true,
        ReadCommentHandling = JsonCommentHandling.Skip,
        AllowTrailingCommas = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.Never,
    };

    /// <summary>
    /// Reads config.json. A missing file gives defaults; a broken one gives defaults plus the error,
    /// so the kiosk still starts and the admin panel can show what went wrong.
    /// </summary>
    public static KioskConfig Load(string path, out string? error)
    {
        error = null;
        if (!File.Exists(path)) return new KioskConfig();
        try
        {
            var text = File.ReadAllText(path);
            var config = JsonSerializer.Deserialize<KioskConfig>(text, Json) ?? new KioskConfig();
            config.Brand ??= new BrandConfig();
            config.Timers ??= new TimerConfig();
            return config;
        }
        catch (Exception ex) when (ex is JsonException or IOException or UnauthorizedAccessException)
        {
            error = $"config.json could not be read: {ex.Message}";
            return new KioskConfig();
        }
    }

    /// <summary>Writes through a temp file so a power cut never leaves half a config.</summary>
    public static void Save(string path, KioskConfig config)
    {
        var dir = Path.GetDirectoryName(Path.GetFullPath(path))!;
        Directory.CreateDirectory(dir);
        var temp = Path.Combine(dir, $".config.{Guid.NewGuid():N}.tmp");
        File.WriteAllText(temp, JsonSerializer.Serialize(config, Json));
        File.Move(temp, path, overwrite: true);
    }
}
