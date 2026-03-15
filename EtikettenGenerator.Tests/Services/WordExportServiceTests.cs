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
    /// Creates a minimal in-memory Word document that mirrors the real template structure:
    /// one outer table with rows×cols cells, each cell containing one inner table with a
    /// paragraph of MERGEFIELD fields (Nachname, Vorname, Medizin, Taktik, RD_Fortbildung,
    /// FS_1, FS_2, FS_3).
    /// </summary>
    private static byte[] CreateTestTemplate(int rows = 10, int cols = 4)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document();
            var body = mainPart.Document.AppendChild(new Body());

            var outerTable = body.AppendChild(new Table());
            for (var r = 0; r < rows; r++)
            {
                var row = outerTable.AppendChild(new TableRow());
                for (var c = 0; c < cols; c++)
                {
                    var outerCell = row.AppendChild(new TableCell());
                    var innerTable = outerCell.AppendChild(new Table());
                    var innerRow = innerTable.AppendChild(new TableRow());
                    var innerCell = innerRow.AppendChild(new TableCell());

                    // Add one paragraph per field so the state machine can traverse them
                    foreach (var fieldName in new[] { "Nachname", "Vorname", "Medizin", "Taktik", "RD_Fortbildung", "FS_1", "FS_2", "FS_3" })
                    {
                        innerCell.AppendChild(BuildMergeFieldParagraph(fieldName));
                    }
                }
            }

            mainPart.Document.Save();
        }

        return ms.ToArray();
    }

    private static Paragraph BuildMergeFieldParagraph(string fieldName)
    {
        var para = new Paragraph();
        para.AppendChild(new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }));
        para.AppendChild(new Run(new FieldCode($" MERGEFIELD {fieldName} ")));
        para.AppendChild(new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }));
        para.AppendChild(new Run(new Text($"«{fieldName}»") { Space = SpaceProcessingModeValues.Preserve }));
        para.AppendChild(new Run(new FieldChar { FieldCharType = FieldCharValues.End }));
        return para;
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
    public void Export_MemberAtPosition1_FieldsReplaced()
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
        allText.Should().NotContain("«Nachname»");
        allText.Should().NotContain("«Vorname»");
    }

    [Fact]
    public void Export_RdFortbildungTrue_ShowsRdf()
    {
        var service = CreateService();
        var member = CreateMember(rdFortbildung: true);
        var template = CreateTestTemplate();
        var pairs = new List<(int, Member)> { (1, member) };

        var bytes = service.ExportWithTemplate(pairs, template);

        using var stream = new MemoryStream(bytes);
        using var doc = WordprocessingDocument.Open(stream, isEditable: false);
        var allText = doc.MainDocumentPart!.Document!.Body!.InnerText;

        allText.Should().Contain("RDF");
    }

    [Fact]
    public void Export_RdFortbildungFalse_ShowsEmpty()
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
        allText.Should().NotContain("«RD_Fortbildung»");
    }

    [Fact]
    public void Export_EmptyPositions_ClearsMergeFieldDisplayValues()
    {
        var service = CreateService();
        var template = CreateTestTemplate();
        // Only fill position 1, leave all others empty
        var pairs = new List<(int, Member)> { (1, CreateMember()) };

        var bytes = service.ExportWithTemplate(pairs, template);

        using var stream = new MemoryStream(bytes);
        using var doc = WordprocessingDocument.Open(stream, isEditable: false);
        var allText = doc.MainDocumentPart!.Document!.Body!.InnerText;

        // No «FieldName» merge display values should remain
        allText.Should().NotContain("«");
        allText.Should().NotContain("»");
    }

    [Fact]
    public void Export_OverflowMembers_ProducesSecondPage()
    {
        var service = CreateService();
        var template = CreateTestTemplate();
        var pairs = Enumerable.Range(1, 45)
            .Select(i => (i, CreateMember($"Name{i}")))
            .ToList<(int, Member)>();

        var bytes = service.ExportWithTemplate(pairs, template);

        using var stream = new MemoryStream(bytes);
        using var doc = WordprocessingDocument.Open(stream, isEditable: false);
        var outerTables = doc.MainDocumentPart!.Document!.Body!.Elements<Table>().ToList();

        // Two outer tables (two pages)
        outerTables.Should().HaveCount(2);
    }

    // --- MapMedQualifikation tests ---

    [Theory]
    [InlineData("Rettungssanitäter/in", "RS")]
    [InlineData("Notarzt / Notärztin",  "NA")]
    [InlineData("Erste-Hilfe",          "EH")]
    [InlineData("",                     "-")]
    [InlineData("Unbekannt",            "-")]
    public void MapMedQualifikation_MapsToAbbreviation(string input, string expected)
        => WordExportService.MapMedQualifikation(input).Should().Be(expected);

    // --- MapDienststellung tests ---

    [Theory]
    [InlineData("Gruppenführer:in",        "GF")]
    [InlineData("Zugführer:in",            "ZF")]
    [InlineData("Helfer:in in Ausbildung", "HF")]
    [InlineData("",                        "HF")]
    [InlineData("Unbekannt",               "HF")]
    [InlineData("ZF mit Stabsausbildung",  "GdSA")]
    public void MapDienststellung_MapsToAbbreviation(string input, string expected)
        => WordExportService.MapDienststellung(input).Should().Be(expected);

    // --- ParseFahrerlaubnis tests ---

    [Theory]
    [InlineData("B", "B", "", "")]
    [InlineData("C", "C", "", "")]
    [InlineData("C1", "C1", "", "")]
    [InlineData("B, C", "C", "", "")]          // C > B in hierarchy
    [InlineData("BE", "", "E", "")]             // E-variant → FS_2
    [InlineData("B BE", "B", "E", "")]          // B + E-variant
    [InlineData("C CE", "C", "E", "")]
    [InlineData("B+E", "", "E", "")]
    [InlineData("A", "", "", "A")]
    [InlineData("A2", "", "", "A2")]
    [InlineData("A A2", "", "", "A")]           // A > A2
    [InlineData("AM", "", "", "")]              // AM excluded
    [InlineData("B A2", "B", "", "A2")]
    [InlineData("C BE A", "C", "E", "A")]
    [InlineData("", "", "", "")]
    [InlineData("  ", "", "", "")]
    public void ParseFahrerlaubnis_VariousInputs(string input,
        string expectedFs1, string expectedFs2, string expectedFs3)
    {
        var (fs1, fs2, fs3) = WordExportService.ParseFahrerlaubnis(input);

        fs1.Should().Be(expectedFs1);
        fs2.Should().Be(expectedFs2);
        fs3.Should().Be(expectedFs3);
    }
}
