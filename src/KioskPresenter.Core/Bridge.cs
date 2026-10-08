using System.Text.Json;
using System.Text.Json.Serialization;

namespace KioskPresenter.Core;

/// <summary>
/// The JSON protocol between ui/index.html and the host. Message names match the Claude Design
/// mock (ui/mock-data.js) plus the host additions documented in the build plan.
/// </summary>
public static class Bridge
{
    public static readonly JsonSerializerOptions Json = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
    };

    /// <summary>Parses a message the UI posted with chrome.webview.postMessage (WebMessageAsJson).</summary>
    public static UiCommand? ParseCommand(string json)
    {
        try
        {
            var cmd = JsonSerializer.Deserialize<UiCommand>(json, Json);
            return cmd is { Cmd.Length: > 0 } ? cmd : null;
        }
        catch (JsonException)
        {
            return null;
        }
    }

    public static string Serialize(object message) => JsonSerializer.Serialize(message, message.GetType(), Json);

    /// <summary>Library message; thumbnails are URLs the WebView can load.</summary>
    public static object Library(LibraryScan scan, Func<DeckInfo, string?> thumbnailUrl) => new
    {
        type = "library",
        status = scan.Status switch
        {
            LibraryStatus.Ok => "ok",
            LibraryStatus.Empty => "empty",
            _ => "unavailable",
        },
        decks = scan.Decks.Select(d => new DeckDto(
            d.Id,
            d.Title,
            d.Description,
            d.VisibleSlideCount ?? 0,
            d.HasVideo,
            thumbnailUrl(d))).ToList(),
    };

    public static object Opening(string id) => new { type = "opening", id };

    public static object ShowStarted(string id, int slideCount, bool video) => new { type = "showStarted", id, slideCount, video };

    public static object Slide(int n, int total, bool video) => new { type = "slide", n, total, video };

    public static object ShowEnded() => new { type = "showEnded" };

    public static object Home() => new { type = "home" };

    public static object OpenFailed(string id, string reason) => new { type = "openFailed", id, reason };

    public static object Activity() => new { type = "activity" };

    public static object Pin() => new { type = "pin" };

    public static object PinResult(bool ok) => new { type = "pinResult", ok };

    public static object SetupResult(bool ok) => new { type = "setupResult", ok };

    public static object AdminStatus(AdminStatusDto status) => new { type = "adminStatus", status };
}

public sealed record UiCommand(string Cmd, string? Id = null, string? Digits = null, string? Action = null, bool? Visible = null);

public sealed record DeckDto(string Id, string Title, string? Description, int SlideCount, bool HasVideo, string? Thumbnail);

/// <summary>Shape the admin panel renders (see ui/mock-data.js STATUS_OK).</summary>
public sealed record AdminStatusDto(
    string Version,
    string Folder,
    bool FolderOk,
    int Decks,
    string? Ppt,
    bool Network,
    string LastRefresh,
    string Account,
    bool Autostart,
    IReadOnlyList<AdminWarning> Warnings);

public sealed record AdminWarning(string Id, string Title, string Detail);
