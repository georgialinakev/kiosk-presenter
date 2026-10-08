namespace KioskPresenter.Core.Tests;

public sealed class ShowMapTests
{
    [Fact]
    public void Position_skips_hidden_slides()
    {
        var map = new ShowMap(5, hiddenSlides: [2, 4], videoSlides: [3]);

        Assert.Equal(3, map.VisibleTotal(fallback: 99));
        Assert.Equal(1, map.Position(1));
        Assert.Equal(2, map.Position(3));
        Assert.Equal(3, map.Position(5));
        Assert.Equal(1, map.Position(2)); // hidden slide reached by a link
        Assert.True(map.IsVideo(3));
        Assert.False(map.IsVideo(1));
    }

    [Fact]
    public void Unknown_total_falls_back_to_powerpoint_count()
    {
        var map = new ShowMap(null, [], []);

        Assert.Equal(12, map.VisibleTotal(fallback: 12));
        Assert.Equal(4, map.Position(4));
    }
}

public sealed class PinHasherTests
{
    [Fact]
    public void Hash_verifies_only_the_same_pin()
    {
        var stored = PinHasher.Hash("2468");

        Assert.StartsWith("pbkdf2-sha256$", stored);
        Assert.DoesNotContain("2468", stored);
        Assert.True(PinHasher.Verify("2468", stored));
        Assert.False(PinHasher.Verify("2469", stored));
        Assert.False(PinHasher.Verify(null, stored));
    }

    [Fact]
    public void Same_pin_hashes_differently_each_time() =>
        Assert.NotEqual(PinHasher.Hash("1234"), PinHasher.Hash("1234"));

    [Theory]
    [InlineData("123", false)]
    [InlineData("1234", true)]
    [InlineData("12345678", true)]
    [InlineData("123456789", false)]
    [InlineData("12a4", false)]
    [InlineData("", false)]
    public void Pin_rules_are_four_to_eight_digits(string pin, bool valid) =>
        Assert.Equal(valid, PinHasher.IsValidPin(pin));

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("plain")]
    [InlineData("pbkdf2-sha256$x$abc$def")]
    [InlineData("pbkdf2-sha256$1000$!!!$###")]
    public void Malformed_hashes_never_verify(string? stored) =>
        Assert.False(PinHasher.Verify("2468", stored));
}

public sealed class ConfigStoreTests : IDisposable
{
    private readonly TestDecks _dir = new();

    public void Dispose() => _dir.Dispose();

    [Fact]
    public void Missing_file_gives_defaults_and_needs_setup()
    {
        var config = ConfigStore.Load(Path.Combine(_dir.Folder, "config.json"), out var error);

        Assert.Null(error);
        Assert.True(config.NeedsSetup);
        Assert.Equal(180, config.Timers.PlayerIdle);
    }

    [Fact]
    public void Save_then_load_round_trips()
    {
        var path = Path.Combine(_dir.Folder, "sub", "config.json");
        var config = new KioskConfig { ContentFolder = @"D:\Decks", PinHash = "x", Brand = { Title = "Hall of Discovery" } };

        ConfigStore.Save(path, config);
        var loaded = ConfigStore.Load(path, out var error);

        Assert.Null(error);
        Assert.Equal(@"D:\Decks", loaded.ContentFolder);
        Assert.Equal("Hall of Discovery", loaded.Brand.Title);
        Assert.False(loaded.NeedsSetup);
        Assert.Contains("\"contentFolder\"", File.ReadAllText(path));
    }

    [Fact]
    public void Broken_file_gives_defaults_and_an_error()
    {
        var path = _dir.Raw("config.json", "{ not json"u8.ToArray());

        var config = ConfigStore.Load(path, out var error);

        Assert.NotNull(error);
        Assert.True(config.NeedsSetup);
    }

    [Fact]
    public void Content_folder_defaults_under_the_data_root()
    {
        var paths = new KioskPaths(_dir.Folder);

        Assert.Equal(paths.DefaultContentFolder, new KioskConfig().ResolveContentFolder(paths));
    }
}

public sealed class BridgeTests
{
    [Fact]
    public void Commands_parse_from_ui_json()
    {
        Assert.Equal(new UiCommand("open", Id: "a1b2"), Bridge.ParseCommand("{\"cmd\":\"open\",\"id\":\"a1b2\"}"));
        Assert.Equal(new UiCommand("prompt", Visible: true), Bridge.ParseCommand("{\"cmd\":\"prompt\",\"visible\":true}"));
        Assert.Equal("exit", Bridge.ParseCommand("{\"cmd\":\"admin\",\"action\":\"exit\"}")!.Action);
        Assert.Equal("2468", Bridge.ParseCommand("{\"cmd\":\"pinSubmit\",\"digits\":\"2468\"}")!.Digits);
    }

    [Theory]
    [InlineData("")]
    [InlineData("not json")]
    [InlineData("{}")]
    [InlineData("{\"cmd\":\"\"}")]
    public void Junk_is_ignored(string json) => Assert.Null(Bridge.ParseCommand(json));

    [Fact]
    public void Library_message_matches_the_ui_shape()
    {
        var deck = new DeckInfo
        {
            Id = "abc", FilePath = @"C:\x\a.pptx", Title = "A", Format = DeckFormat.OpenXml,
            TotalSlides = 3, HiddenSlides = [1], VideoSlides = [2], LastWriteUtc = DateTime.UtcNow, Length = 10,
        };
        var scan = new LibraryScan(LibraryStatus.Ok, [deck], [], @"C:\x");

        var json = Bridge.Serialize(Bridge.Library(scan, _ => "https://cache.kiosk.local/thumbs/abc.png"));

        Assert.Contains("\"type\":\"library\"", json);
        Assert.Contains("\"status\":\"ok\"", json);
        Assert.Contains("\"slideCount\":2", json);
        Assert.Contains("\"hasVideo\":true", json);
        Assert.Contains("\"thumbnail\":\"https://cache.kiosk.local/thumbs/abc.png\"", json);
        Assert.DoesNotContain("filePath", json, StringComparison.OrdinalIgnoreCase);
    }
}

public sealed class UiConfigScriptTests
{
    [Fact]
    public void Script_turns_dev_off_and_follows_pin_state()
    {
        var script = UiConfigScript.Build(new KioskConfig(), logoUrl: null);

        Assert.StartsWith("// Generated", script);
        Assert.Contains("window.KIOSK_HOST_CONFIG = {", script);
        Assert.Contains("\"dev\": false", script);
        Assert.Contains("\"firstRun\": true", script);
        Assert.Contains("\"playerIdle\": 180", script);

        var withPin = UiConfigScript.Build(new KioskConfig { PinHash = "x" }, null);
        Assert.Contains("\"firstRun\": false", withPin);
    }

    [Theory]
    [InlineData("#4fd1b5", true)]
    [InlineData("#4FD1B5", true)]
    [InlineData("4fd1b5", false)]
    [InlineData("#4fd1b", false)]
    [InlineData("red", false)]
    [InlineData(null, false)]
    public void Accent_must_be_a_hex_colour(string? accent, bool ok) =>
        Assert.Equal(ok, UiConfigScript.IsHexColour(accent));
}
