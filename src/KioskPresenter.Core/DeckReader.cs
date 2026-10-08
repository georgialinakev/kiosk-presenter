using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;

namespace KioskPresenter.Core;

/// <summary>
/// Reads what the library needs from a presentation without PowerPoint: slide order, hidden
/// slides, which slides hold video, and the embedded preview image. .pptx/.ppsx are zip files
/// of XML, so the BCL zip and XML readers are enough.
/// </summary>
public static partial class DeckReader
{
    public static readonly string[] Extensions = [".pptx", ".ppsx", ".ppt"];

    private static readonly XNamespace P = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace Rel = "http://schemas.openxmlformats.org/package/2006/relationships";
    private static readonly XNamespace Dc = "http://purl.org/dc/elements/1.1/";
    private const string VideoRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/video";
    private const string ThumbnailRelType = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail";

    private static readonly byte[] ZipMagic = [0x50, 0x4B, 0x03, 0x04];
    private static readonly byte[] CfbMagic = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

    public static bool IsCandidate(string path)
    {
        var name = Path.GetFileName(path);
        if (name.StartsWith("~$", StringComparison.Ordinal)) return false; // Office lock file
        return Extensions.Contains(Path.GetExtension(path), StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>Reads one deck. Throws <see cref="DeckReadException"/> with a staff-readable reason when it can't be listed.</summary>
    public static DeckInfo Read(string path)
    {
        var full = Path.GetFullPath(path);
        var info = new FileInfo(full);
        if (!info.Exists) throw new DeckReadException("File not found.");
        if (info.Length == 0) throw new DeckReadException("File is empty.");

        var header = ReadHeader(full);
        var extension = info.Extension.ToLowerInvariant();
        var common = new
        {
            Id = MakeId(full),
            Title = TitleFromFileName(full),
            info.LastWriteTimeUtc,
            info.Length,
        };

        if (header.AsSpan().StartsWith(CfbMagic))
        {
            // An Open XML file wrapped in a compound file is how PowerPoint stores password-protected decks.
            if (extension != ".ppt") throw new DeckReadException("Password-protected presentations can't be played on the kiosk.");
            return new DeckInfo
            {
                Id = common.Id,
                FilePath = full,
                Title = common.Title,
                Format = DeckFormat.Legacy,
                LastWriteUtc = common.LastWriteTimeUtc,
                Length = common.Length,
            };
        }

        if (!header.AsSpan().StartsWith(ZipMagic)) throw new DeckReadException("Not a PowerPoint file.");

        try
        {
            using var zip = ZipFile.OpenRead(full);
            var slides = ReadSlideOrder(zip);
            var hidden = new List<int>();
            var video = new List<int>();
            for (var i = 0; i < slides.Count; i++)
            {
                var (isHidden, hasVideo) = InspectSlide(zip, slides[i]);
                if (isHidden) hidden.Add(i + 1);
                if (hasVideo) video.Add(i + 1);
            }

            return new DeckInfo
            {
                Id = common.Id,
                FilePath = full,
                Title = common.Title,
                Description = ReadDescription(zip),
                Format = DeckFormat.OpenXml,
                TotalSlides = slides.Count,
                HiddenSlides = hidden,
                VideoSlides = video,
                LastWriteUtc = common.LastWriteTimeUtc,
                Length = common.Length,
            };
        }
        catch (InvalidDataException)
        {
            throw new DeckReadException("File is damaged or not a PowerPoint file.");
        }
        catch (XmlException)
        {
            throw new DeckReadException("File is damaged (bad slide XML).");
        }
    }

    /// <summary>
    /// Copies the preview PowerPoint stores inside the file (docProps/thumbnail.jpeg) to <paramref name="destination"/>.
    /// Returns false when the deck has none.
    /// </summary>
    public static bool TryExtractThumbnail(string path, string destination)
    {
        try
        {
            using var zip = ZipFile.OpenRead(path);
            var target = FindPackageRelationship(zip, ThumbnailRelType) ?? "docProps/thumbnail.jpeg";
            var entry = zip.GetEntry(target);
            if (entry is null || entry.Length == 0) return false;
            Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(destination))!);
            entry.ExtractToFile(destination, overwrite: true);
            return true;
        }
        catch (Exception ex) when (ex is IOException or InvalidDataException or UnauthorizedAccessException or XmlException)
        {
            return false;
        }
    }

    /// <summary>Display title from the file name: "03_Age-of-Steam.pptx" becomes "Age of Steam".</summary>
    /// <remarks>A leading number lets staff control order by file name without it showing on screen.</remarks>
    public static string TitleFromFileName(string path)
    {
        var stem = Path.GetFileNameWithoutExtension(path);
        var title = OrderPrefix().Replace(stem, string.Empty);
        title = title.Replace('_', ' ');
        title = Spaces().Replace(title, " ").Trim();
        return title.Length > 0 ? title : stem.Trim();
    }

    /// <summary>Ordering key: file name, case-insensitive, with numbers compared as numbers.</summary>
    public static string SortKey(string path) =>
        Digits().Replace(Path.GetFileName(path).ToUpperInvariant(), m => m.Value.PadLeft(10, '0'));

    public static string MakeId(string fullPath)
    {
        var bytes = SHA256.HashData(Encoding.UTF8.GetBytes(Path.GetFullPath(fullPath).ToUpperInvariant()));
        return Convert.ToHexStringLower(bytes)[..12];
    }

    private static byte[] ReadHeader(string path)
    {
        try
        {
            using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
            var buffer = new byte[8];
            var read = stream.ReadAtLeast(buffer, buffer.Length, throwOnEndOfStream: false);
            return buffer[..read];
        }
        catch (IOException ex)
        {
            throw new DeckReadException($"File is in use or still copying ({ex.Message.TrimEnd('.')}).");
        }
        catch (UnauthorizedAccessException)
        {
            throw new DeckReadException("The kiosk account can't read this file.");
        }
    }

    /// <summary>Slide part paths in presentation order, from presentation.xml's slide list.</summary>
    private static List<string> ReadSlideOrder(ZipArchive zip)
    {
        var presentation = LoadXml(zip, "ppt/presentation.xml") ?? throw new DeckReadException("Not a PowerPoint presentation.");
        var rels = LoadRelationships(zip, "ppt/presentation.xml");
        var result = new List<string>();
        var list = presentation.Root?.Element(P + "sldIdLst");
        if (list is null) return result;
        foreach (var slideId in list.Elements(P + "sldId"))
        {
            var rid = (string?)slideId.Attribute(R + "id");
            if (rid is not null && rels.TryGetValue(rid, out var rel)) result.Add(rel.Target);
        }
        return result;
    }

    private static (bool Hidden, bool Video) InspectSlide(ZipArchive zip, string slidePath)
    {
        var slide = LoadXml(zip, slidePath);
        if (slide?.Root is null) return (false, false);
        var hidden = (string?)slide.Root.Attribute("show") is "0" or "false";
        var video = slide.Descendants(A + "videoFile").Any()
            || LoadRelationships(zip, slidePath).Values.Any(r => r.Type == VideoRelType);
        return (hidden, video);
    }

    private static string? ReadDescription(ZipArchive zip)
    {
        var core = LoadXml(zip, "docProps/core.xml");
        var text = core?.Root?.Element(Dc + "description")?.Value?.Trim();
        return string.IsNullOrEmpty(text) ? null : text;
    }

    private static string? FindPackageRelationship(ZipArchive zip, string type)
    {
        var rels = LoadXml(zip, "_rels/.rels");
        var target = rels?.Root?.Elements(Rel + "Relationship")
            .FirstOrDefault(e => (string?)e.Attribute("Type") == type)?
            .Attribute("Target")?.Value;
        return target is null ? null : ResolvePartPath(string.Empty, target);
    }

    private sealed record Relationship(string Type, string Target);

    /// <summary>Relationships of a part, with targets resolved to zip entry paths.</summary>
    private static Dictionary<string, Relationship> LoadRelationships(ZipArchive zip, string partPath)
    {
        var dir = partPath.Contains('/') ? partPath[..partPath.LastIndexOf('/')] : string.Empty;
        var name = partPath[(partPath.LastIndexOf('/') + 1)..];
        var relsPath = (dir.Length > 0 ? dir + "/" : string.Empty) + "_rels/" + name + ".rels";
        var result = new Dictionary<string, Relationship>(StringComparer.Ordinal);
        var doc = LoadXml(zip, relsPath);
        if (doc?.Root is null) return result;
        foreach (var e in doc.Root.Elements(Rel + "Relationship"))
        {
            var id = (string?)e.Attribute("Id");
            var type = (string?)e.Attribute("Type") ?? string.Empty;
            var target = (string?)e.Attribute("Target") ?? string.Empty;
            var external = string.Equals((string?)e.Attribute("TargetMode"), "External", StringComparison.OrdinalIgnoreCase);
            if (id is null) continue;
            result[id] = new Relationship(type, external ? target : ResolvePartPath(dir, target));
        }
        return result;
    }

    /// <summary>Resolves a relationship target against the source part's folder ("slides/slide1.xml", "../media/x.png", "/ppt/x").</summary>
    internal static string ResolvePartPath(string sourceDir, string target)
    {
        if (target.StartsWith('/')) return target.TrimStart('/');
        var parts = new List<string>(sourceDir.Split('/', StringSplitOptions.RemoveEmptyEntries));
        foreach (var segment in target.Split('/', StringSplitOptions.RemoveEmptyEntries))
        {
            if (segment == ".") continue;
            if (segment == "..") { if (parts.Count > 0) parts.RemoveAt(parts.Count - 1); continue; }
            parts.Add(segment);
        }
        return string.Join('/', parts);
    }

    private static XDocument? LoadXml(ZipArchive zip, string path)
    {
        var entry = zip.GetEntry(path);
        if (entry is null) return null;
        using var stream = entry.Open();
        var settings = new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null };
        using var reader = XmlReader.Create(stream, settings);
        return XDocument.Load(reader);
    }

    [GeneratedRegex(@"^\d{1,4}\s*[-_.)]+\s*")]
    private static partial Regex OrderPrefix();

    [GeneratedRegex(@"\s+")]
    private static partial Regex Spaces();

    [GeneratedRegex(@"\d+")]
    private static partial Regex Digits();
}
