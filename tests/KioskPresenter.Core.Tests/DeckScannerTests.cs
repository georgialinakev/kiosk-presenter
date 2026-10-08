namespace KioskPresenter.Core.Tests;

public sealed class DeckScannerTests : IDisposable
{
    private readonly TestDecks _decks = new();

    public void Dispose() => _decks.Dispose();

    [Fact]
    public void Missing_folder_is_unavailable()
    {
        var scan = DeckScanner.Scan(Path.Combine(_decks.Folder, "nope"));

        Assert.Equal(LibraryStatus.Unavailable, scan.Status);
        Assert.Empty(scan.Decks);
        Assert.NotNull(scan.Error);
    }

    [Fact]
    public void Folder_without_decks_is_empty()
    {
        _decks.Raw("readme.txt", [1, 2, 3]);

        Assert.Equal(LibraryStatus.Empty, DeckScanner.Scan(_decks.Folder).Status);
    }

    [Fact]
    public void Scan_skips_lock_files_reports_bad_files_and_orders_by_name()
    {
        _decks.Pptx("10_Later.pptx", [new()]);
        _decks.Pptx("2_Earlier.pptx", [new(), new()]);
        _decks.Raw("~$2_Earlier.pptx", [1, 2, 3]);
        _decks.Raw("Locked.pptx", TestDecks.CompoundFileHeader);

        var scan = DeckScanner.Scan(_decks.Folder);

        Assert.Equal(LibraryStatus.Ok, scan.Status);
        Assert.Equal(new[] { "Earlier", "Later" }, scan.Decks.Select(d => d.Title).ToArray());
        var issue = Assert.Single(scan.Issues);
        Assert.EndsWith("Locked.pptx", issue.FilePath);
        Assert.NotNull(scan.Find(scan.Decks[0].Id));
        Assert.Null(scan.Find("missing"));
    }
}
