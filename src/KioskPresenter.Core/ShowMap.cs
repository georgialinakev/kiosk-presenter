namespace KioskPresenter.Core;

/// <summary>
/// Turns PowerPoint's absolute slide index into what the visitor sees: position among visible
/// slides, the visible total, and whether the slide holds video (for the idle exemption).
/// </summary>
public sealed class ShowMap
{
    private readonly int? _total;
    private readonly HashSet<int> _hidden;
    private readonly HashSet<int> _video;

    public ShowMap(int? totalSlides, IEnumerable<int> hiddenSlides, IEnumerable<int> videoSlides)
    {
        _total = totalSlides;
        _hidden = new HashSet<int>(hiddenSlides);
        _video = new HashSet<int>(videoSlides);
    }

    /// <summary>Visible slide count, or the fallback PowerPoint reports when the file couldn't be read up front.</summary>
    public int VisibleTotal(int fallback) =>
        _total is int total ? Math.Max(1, total - _hidden.Count(i => i >= 1 && i <= total)) : Math.Max(1, fallback);

    /// <summary>1-based position among visible slides. A hidden slide (reached by a link) maps to the nearest visible one before it.</summary>
    public int Position(int slideIndex)
    {
        if (slideIndex < 1) return 1;
        var position = 0;
        for (var i = 1; i <= slideIndex; i++)
        {
            if (!_hidden.Contains(i)) position++;
        }
        return Math.Max(1, position);
    }

    public bool IsVideo(int slideIndex) => _video.Contains(slideIndex);
}
