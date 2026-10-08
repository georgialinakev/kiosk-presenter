namespace KioskPresenter.Core;

/// <summary>Builds the library from the content folder (top level only).</summary>
public static class DeckScanner
{
    public static LibraryScan Scan(string folder)
    {
        string full;
        try
        {
            full = Path.GetFullPath(folder);
        }
        catch (Exception ex) when (ex is ArgumentException or NotSupportedException or PathTooLongException)
        {
            return new LibraryScan(LibraryStatus.Unavailable, [], [], folder, $"Content folder path is invalid: {ex.Message}");
        }

        if (!Directory.Exists(full))
        {
            return new LibraryScan(LibraryStatus.Unavailable, [], [], full, "Content folder does not exist.");
        }

        IEnumerable<string> files;
        try
        {
            files = Directory.EnumerateFiles(full, "*", SearchOption.TopDirectoryOnly)
                .Where(DeckReader.IsCandidate)
                .Where(f => !IsHiddenOrSystem(f))
                .OrderBy(DeckReader.SortKey, StringComparer.Ordinal)
                .ToList();
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            return new LibraryScan(LibraryStatus.Unavailable, [], [], full, $"Content folder can't be read: {ex.Message}");
        }

        var decks = new List<DeckInfo>();
        var issues = new List<ScanIssue>();
        foreach (var file in files)
        {
            try
            {
                decks.Add(DeckReader.Read(file));
            }
            catch (DeckReadException ex)
            {
                issues.Add(new ScanIssue(file, ex.Message));
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                issues.Add(new ScanIssue(file, ex.Message));
            }
        }

        return new LibraryScan(decks.Count > 0 ? LibraryStatus.Ok : LibraryStatus.Empty, decks, issues, full);
    }

    private static bool IsHiddenOrSystem(string path)
    {
        try
        {
            var attributes = File.GetAttributes(path);
            return (attributes & (FileAttributes.Hidden | FileAttributes.System)) != 0;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            return false;
        }
    }
}
