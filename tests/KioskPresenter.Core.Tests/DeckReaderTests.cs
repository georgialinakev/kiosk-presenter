namespace KioskPresenter.Core.Tests;

public sealed class DeckReaderTests : IDisposable
{
    private readonly TestDecks _decks = new();

    public void Dispose() => _decks.Dispose();

    [Fact]
    public void Demo_gallery_tour_has_six_visible_slides_and_no_video()
    {
        var deck = DeckReader.Read(TestDecks.Fixture("Kiosk_Demo_Gallery_Tour.pptx"));

        Assert.Equal(DeckFormat.OpenXml, deck.Format);
        Assert.Equal("Kiosk Demo Gallery Tour", deck.Title);
        Assert.Equal(6, deck.TotalSlides);
        Assert.Equal(6, deck.VisibleSlideCount);
        Assert.Empty(deck.HiddenSlides);
        Assert.False(deck.HasVideo);
    }

    [Fact]
    public void Demo_video_spotlight_has_video_on_slide_two()
    {
        var deck = DeckReader.Read(TestDecks.Fixture("Kiosk_Demo_Video_Spotlight.pptx"));

        Assert.Equal(4, deck.VisibleSlideCount);
        Assert.Equal(new[] { 2 }, deck.VideoSlides);
        Assert.True(deck.HasVideo);
    }

    [Fact]
    public void Hidden_slides_are_counted_out()
    {
        var path = _decks.Pptx("hidden.pptx", [new(), new(Hidden: true), new()]);

        var deck = DeckReader.Read(path);

        Assert.Equal(3, deck.TotalSlides);
        Assert.Equal(2, deck.VisibleSlideCount);
        Assert.Equal(new[] { 2 }, deck.HiddenSlides);
    }

    [Fact]
    public void Video_is_found_from_the_slide_relationship_alone()
    {
        var path = _decks.Pptx("relonly.pptx", [new(), new(), new(VideoRelOnly: true)]);

        Assert.Equal(new[] { 3 }, DeckReader.Read(path).VideoSlides);
    }

    [Fact]
    public void Description_comes_from_document_properties()
    {
        var path = _decks.Pptx("desc.pptx", [new()], description: "Farming, trade and belief");

        Assert.Equal("Farming, trade and belief", DeckReader.Read(path).Description);
    }

    [Fact]
    public void Password_protected_pptx_is_rejected_with_a_reason()
    {
        var path = _decks.Raw("secret.pptx", TestDecks.CompoundFileHeader);

        var ex = Assert.Throws<DeckReadException>(() => DeckReader.Read(path));
        Assert.Contains("Password", ex.Message);
    }

    [Fact]
    public void Legacy_ppt_is_listed_without_slide_details()
    {
        var path = _decks.Raw("old_deck.ppt", TestDecks.CompoundFileHeader);

        var deck = DeckReader.Read(path);

        Assert.Equal(DeckFormat.Legacy, deck.Format);
        Assert.Null(deck.TotalSlides);
        Assert.Null(deck.VisibleSlideCount);
        Assert.Equal("old deck", deck.Title);
    }

    [Theory]
    [InlineData("empty.pptx", new byte[0])]
    [InlineData("garbage.pptx", new byte[] { 1, 2, 3, 4, 5, 6, 7, 8 })]
    [InlineData("broken.pptx", new byte[] { 0x50, 0x4B, 0x03, 0x04, 9, 9, 9, 9, 9, 9 })]
    public void Unreadable_files_are_rejected(string name, byte[] bytes)
    {
        var path = _decks.Raw(name, bytes);

        Assert.Throws<DeckReadException>(() => DeckReader.Read(path));
    }

    [Fact]
    public void Zip_that_is_not_a_presentation_is_rejected()
    {
        var path = Path.Combine(_decks.Folder, "notes.pptx");
        using (var zip = System.IO.Compression.ZipFile.Open(path, System.IO.Compression.ZipArchiveMode.Create))
        {
            zip.CreateEntry("word/document.xml");
        }

        Assert.Throws<DeckReadException>(() => DeckReader.Read(path));
    }

    [Theory]
    [InlineData("Kiosk_Demo_Gallery_Tour.pptx", "Kiosk Demo Gallery Tour")]
    [InlineData("03_Age-of-Steam.pptx", "Age-of-Steam")]
    [InlineData("12 - Masks of West Africa.ppsx", "Masks of West Africa")]
    [InlineData("1955 Atomic Clocks.pptx", "1955 Atomic Clocks")]
    [InlineData("01.pptx", "01")]
    [InlineData("  spaced   out  .pptx", "spaced out")]
    public void Title_comes_from_the_file_name(string file, string expected) =>
        Assert.Equal(expected, DeckReader.TitleFromFileName(file));

    [Fact]
    public void Sort_key_orders_numbers_numerically()
    {
        var files = new[] { "10_b.pptx", "2_a.pptx", "1_c.pptx" };

        var ordered = files.OrderBy(DeckReader.SortKey, StringComparer.Ordinal).ToArray();

        Assert.Equal(new[] { "1_c.pptx", "2_a.pptx", "10_b.pptx" }, ordered);
    }

    [Theory]
    [InlineData("ppt", "slides/slide1.xml", "ppt/slides/slide1.xml")]
    [InlineData("ppt/slides", "../media/image1.png", "ppt/media/image1.png")]
    [InlineData("ppt", "/ppt/slides/slide2.xml", "ppt/slides/slide2.xml")]
    [InlineData("", "docProps/thumbnail.jpeg", "docProps/thumbnail.jpeg")]
    public void Relationship_targets_resolve_to_zip_paths(string dir, string target, string expected) =>
        Assert.Equal(expected, DeckReader.ResolvePartPath(dir, target));

    [Fact]
    public void Embedded_thumbnail_is_extracted_when_present()
    {
        var path = _decks.Pptx("thumb.pptx", [new()], thumbnail: true);
        var dest = Path.Combine(_decks.Folder, "out", "t.jpg");

        Assert.True(DeckReader.TryExtractThumbnail(path, dest));
        Assert.True(new FileInfo(dest).Length > 0);
    }

    [Fact]
    public void Missing_thumbnail_returns_false()
    {
        var dest = Path.Combine(_decks.Folder, "none.jpg");

        Assert.False(DeckReader.TryExtractThumbnail(TestDecks.Fixture("Kiosk_Demo_Gallery_Tour.pptx"), dest));
    }

    [Fact]
    public void Ids_are_stable_and_path_specific()
    {
        var a = Path.Combine(_decks.Folder, "a.pptx");
        var b = Path.Combine(_decks.Folder, "b.pptx");

        Assert.Equal(DeckReader.MakeId(a), DeckReader.MakeId(a));
        Assert.NotEqual(DeckReader.MakeId(a), DeckReader.MakeId(b));
        Assert.Equal(12, DeckReader.MakeId(a).Length);
    }

    [Theory]
    [InlineData("~$deck.pptx", false)]
    [InlineData("deck.pptx", true)]
    [InlineData("deck.PPSX", true)]
    [InlineData("deck.ppt", true)]
    [InlineData("deck.pptm", false)]
    [InlineData("deck.pdf", false)]
    public void Only_presentations_are_candidates(string name, bool expected) =>
        Assert.Equal(expected, DeckReader.IsCandidate(name));
}
