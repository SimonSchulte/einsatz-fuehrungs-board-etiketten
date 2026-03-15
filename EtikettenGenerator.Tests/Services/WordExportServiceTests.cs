using AwesomeAssertions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using EtikettenGenerator.Web.Models;
using EtikettenGenerator.Web.Services;
using Microsoft.Extensions.Logging.Abstractions;

namespace EtikettenGenerator.Tests.Services;

public sealed class WordExportServiceTests
{
    private static WordExportService CreateService() =>
        new(NullLogger<WordExportService>.Instance);

    private static Member CreateMember(string nachname = "Mustermann", string vorname = "Max",
        string medQual = "Notarzt", string dienst = "Staffelführer",
        string fahrerlaubnis = "B", bool rdFortbildung = true) =>
        new()
        {
            Nachname = nachname,
            Vorname = vorname,
            MedQualifikation = medQual,
            Dienststellung = dienst,
            Fahrerlaubnis = fahrerlaubnis,
            HatRettungsdienstfortbildung = rdFortbildung,
        };

    /// <summary>
    /// Creates a minimal in-memory Word document with a 10×4 table containing the
    /// standard label placeholders in every cell. This replaces the embedded template
    /// so tests are self-contained and not dependent on the production .docx file.
    /// </summary>
    private static byte[] CreateTestTemplate(int rows = 10, int cols = 4)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document();
            var body = mainPart.Document.AppendChild(new Body());

            var table = body.AppendChild(new Table());
            for (var r = 0; r < rows; r++)
            {
                var row = table.AppendChild(new TableRow());
                for (var c = 0; c < cols; c++)
                {
                    var cell = row.AppendChild(new TableCell());
                    var para = cell.AppendChild(new Paragraph());
                    para.AppendChild(new Run(new Text(
                        "{{NACHNAME}} {{VORNAME}} {{MED_QUAL}} {{DIENSTSTELLUNG}} {{FAHRERLAUBNIS}} {{RD_FORTBILDUNG}}")
                    { Space = SpaceProcessingModeValues.Preserve }));
                }
            }

            mainPart.Document.Save();
        }

        return ms.ToArray();
    }

    [Fact]
    public void Export_SingleMember_ReturnsValidDocx()
    {
        var service = CreateService();
        var member = CreateMember();
        var template = CreateTestTemplate();
        var pairs = new List<(int, Member)> { (1, member) };

        var result = service.ExportWithTemplate(pairs, template);

        result.Should().NotBeEmpty();
        // PK zip magic bytes → valid docx
        result[0].Should().Be((byte)'P');
        result[1].Should().Be((byte)'K');
    }

    [Fact]
    public void Export_MemberAtPosition1_PlaceholdersReplaced()
    {
        var service = CreateService();
        var member = CreateMember("Müller", "Hans", "RS", "Gruppenführer", "C", false);
        var template = CreateTestTemplate();
        var pairs = new List<(int, Member)> { (1, member) };

        var bytes = service.ExportWithTemplate(pairs, template);

        using var stream = new MemoryStream(bytes);
        using var doc = WordprocessingDocument.Open(stream, isEditable: false);
        var allText = doc.MainDocumentPart!.Document!.Body!.InnerText;

        allText.Should().Contain("Müller");
        allText.Should().Contain("Hans");
        allText.Should().NotContain("{{NACHNAME}}");
        allText.Should().NotContain("{{VORNAME}}");
    }

    [Fact]
    public void Export_RdFortbildungTrue_ShowsJa()
    {
        var service = CreateService();
        var member = CreateMember(rdFortbildung: true);
        var template = CreateTestTemplate();
        var pairs = new List<(int, Member)> { (1, member) };

        var bytes = service.ExportWithTemplate(pairs, template);

        using var stream = new MemoryStream(bytes);
        using var doc = WordprocessingDocument.Open(stream, isEditable: false);
        var allText = doc.MainDocumentPart!.Document!.Body!.InnerText;

        allText.Should().Contain("Ja");
    }

    [Fact]
    public void Export_RdFortbildungFalse_ShowsEmptyNotNein()
    {
        var service = CreateService();
        var member = CreateMember(rdFortbildung: false);
        var template = CreateTestTemplate();
        var pairs = new List<(int, Member)> { (1, member) };

        var bytes = service.ExportWithTemplate(pairs, template);

        using var stream = new MemoryStream(bytes);
        using var doc = WordprocessingDocument.Open(stream, isEditable: false);
        var allText = doc.MainDocumentPart!.Document!.Body!.InnerText;

        allText.Should().NotContain("Nein");
        allText.Should().NotContain("{{RD_FORTBILDUNG}}");
    }

    [Fact]
    public void Export_NoPlaceholdersRemainInOutput()
    {
        var service = CreateService();
        var template = CreateTestTemplate();
        var pairs = Enumerable.Range(1, 3)
            .Select(i => (i, CreateMember($"Name{i}")))
            .ToList<(int, Member)>();

        var bytes = service.ExportWithTemplate(pairs, template);

        using var stream = new MemoryStream(bytes);
        using var doc = WordprocessingDocument.Open(stream, isEditable: false);
        var allText = doc.MainDocumentPart!.Document!.Body!.InnerText;

        allText.Should().NotContain("{{");
        allText.Should().NotContain("}}");
    }

    [Fact]
    public void Export_PositionMapping_MemberAtCorrectCell()
    {
        var service = CreateService();
        var memberAt5 = CreateMember("AnFünf");
        var template = CreateTestTemplate();
        var pairs = new List<(int, Member)> { (5, memberAt5) };

        var bytes = service.ExportWithTemplate(pairs, template);

        using var stream = new MemoryStream(bytes);
        using var doc = WordprocessingDocument.Open(stream, isEditable: false);
        var cells = doc.MainDocumentPart!.Document!.Body!
            .Descendants<TableCell>().ToList();

        // Cell at index 4 (position 5) should contain the member name
        cells[4].InnerText.Should().Contain("AnFünf");
        // First cell should not contain it
        cells[0].InnerText.Should().NotContain("AnFünf");
    }

    [Fact]
    public void Export_EmptyPosition_ClearsPlaceholders()
    {
        var service = CreateService();
        var template = CreateTestTemplate();
        // Only fill position 1, leave all others empty
        var pairs = new List<(int, Member)> { (1, CreateMember()) };

        var bytes = service.ExportWithTemplate(pairs, template);

        using var stream = new MemoryStream(bytes);
        using var doc = WordprocessingDocument.Open(stream, isEditable: false);
        var cells = doc.MainDocumentPart!.Document!.Body!
            .Descendants<TableCell>().ToList();

        // Cell 2 (position 2) should have no placeholder text
        cells[1].InnerText.Should().NotContain("{{");
    }
}
