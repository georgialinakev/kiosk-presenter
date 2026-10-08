using System.IO.Compression;
using System.Text;

namespace KioskPresenter.Core.Tests;

/// <summary>Builds minimal .pptx files on disk: just the parts DeckReader reads.</summary>
internal sealed class TestDecks : IDisposable
{
    public string Folder { get; } = Path.Combine(Path.GetTempPath(), "kiosk-tests-" + Guid.NewGuid().ToString("N"));

    public TestDecks() => Directory.CreateDirectory(Folder);

    public static string Fixture(string name) => Path.Combine(AppContext.BaseDirectory, "fixtures", name);

    public sealed record SlideSpec(bool Hidden = false, bool VideoElement = false, bool VideoRelOnly = false);

    public string Pptx(string fileName, IReadOnlyList<SlideSpec> slides, string? description = null, bool thumbnail = false)
    {
        var path = Path.Combine(Folder, fileName);
        using var zip = ZipFile.Open(path, ZipArchiveMode.Create);

        var ids = new StringBuilder();
        var rels = new StringBuilder();
        for (var i = 0; i < slides.Count; i++)
        {
            ids.Append($"<p:sldId id=\"{256 + i}\" r:id=\"rId{i + 10}\"/>");
            rels.Append($"<Relationship Id=\"rId{i + 10}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide\" Target=\"slides/slide{i + 1}.xml\"/>");

            var s = slides[i];
            var show = s.Hidden ? " show=\"0\"" : string.Empty;
            var video = s.VideoElement ? "<p:pic><p:nvPicPr><p:nvPr><a:videoFile r:link=\"rId5\"/></p:nvPr></p:nvPicPr></p:pic>" : string.Empty;
            Write(zip, $"ppt/slides/slide{i + 1}.xml",
                $"<p:sld xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"{show}><p:cSld><p:spTree>{video}</p:spTree></p:cSld></p:sld>");
            if (s.VideoElement || s.VideoRelOnly)
            {
                Write(zip, $"ppt/slides/_rels/slide{i + 1}.xml.rels",
                    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                    + "<Relationship Id=\"rId5\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/video\" Target=\"https://www.youtube.com/embed/x\" TargetMode=\"External\"/>"
                    + "</Relationships>");
            }
        }

        Write(zip, "ppt/presentation.xml",
            "<p:presentation xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
            + $"<p:sldIdLst>{ids}</p:sldIdLst></p:presentation>");
        Write(zip, "ppt/_rels/presentation.xml.rels",
            $"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">{rels}</Relationships>");

        if (description is not null)
        {
            Write(zip, "docProps/core.xml",
                "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\">"
                + $"<dc:description>{description}</dc:description></cp:coreProperties>");
        }

        if (thumbnail)
        {
            Write(zip, "_rels/.rels",
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + "<Relationship Id=\"rId9\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail\" Target=\"docProps/thumbnail.jpeg\"/>"
                + "</Relationships>");
            var entry = zip.CreateEntry("docProps/thumbnail.jpeg");
            using var stream = entry.Open();
            stream.Write(new byte[] { 0xFF, 0xD8, 0xFF, 0xE0, 1, 2, 3, 4 });
        }

        return path;
    }

    public string Raw(string fileName, byte[] bytes)
    {
        var path = Path.Combine(Folder, fileName);
        File.WriteAllBytes(path, bytes);
        return path;
    }

    public static readonly byte[] CompoundFileHeader = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1, 0, 0, 0, 0];

    private static void Write(ZipArchive zip, string name, string xml)
    {
        var entry = zip.CreateEntry(name);
        using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(false));
        writer.Write(xml);
    }

    public void Dispose()
    {
        try { Directory.Delete(Folder, recursive: true); } catch { /* temp files */ }
    }
}
