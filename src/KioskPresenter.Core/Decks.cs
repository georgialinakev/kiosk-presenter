namespace KioskPresenter.Core;

public enum DeckFormat
{
    /// <summary>.pptx / .ppsx: metadata read straight from the zip.</summary>
    OpenXml,

    /// <summary>Binary .ppt: playable, but slide count and video are unknown until PowerPoint opens it.</summary>
    Legacy,
}

/// <summary>One presentation in the content folder.</summary>
public sealed record DeckInfo
{
    /// <summary>Stable across scans and edits: derived from the full path.</summary>
    public required string Id { get; init; }

    public required string FilePath { get; init; }

    public required string Title { get; init; }

    public string? Description { get; init; }

    public required DeckFormat Format { get; init; }

    /// <summary>Total slides in the file, hidden ones included. Null for legacy .ppt.</summary>
    public int? TotalSlides { get; init; }

    /// <summary>1-based indexes (presentation order) of slides marked hidden.</summary>
    public IReadOnlyList<int> HiddenSlides { get; init; } = [];

    /// <summary>1-based indexes (presentation order) of slides holding a video, embedded or online.</summary>
    public IReadOnlyList<int> VideoSlides { get; init; } = [];

    public required DateTime LastWriteUtc { get; init; }

    public required long Length { get; init; }

    /// <summary>Slides a visitor will actually see. Null for legacy .ppt.</summary>
    public int? VisibleSlideCount => TotalSlides is int total ? total - HiddenSlides.Count(i => i >= 1 && i <= total) : null;

    public bool HasVideo => VideoSlides.Count > 0;

    /// <summary>Key for cached artefacts; changes whenever the file does.</summary>
    public string CacheKey => $"{Id}-{LastWriteUtc.Ticks:x}-{Length:x}";

    public ShowMap ToShowMap() => new(TotalSlides, HiddenSlides, VideoSlides);
}

public enum LibraryStatus
{
    Ok,
    Empty,
    Unavailable,
}

public sealed record ScanIssue(string FilePath, string Reason);

public sealed record LibraryScan(LibraryStatus Status, IReadOnlyList<DeckInfo> Decks, IReadOnlyList<ScanIssue> Issues, string Folder, string? Error = null)
{
    public DeckInfo? Find(string? id) => id is null ? null : Decks.FirstOrDefault(d => d.Id == id);
}

/// <summary>Why a file was left out of the library.</summary>
public sealed class DeckReadException(string reason) : Exception(reason);
