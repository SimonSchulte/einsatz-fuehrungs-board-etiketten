using System.Diagnostics;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using EtikettenGenerator.Web.Models;
using Microsoft.Extensions.Logging;

namespace EtikettenGenerator.Web.Services;

public sealed class WordExportService(ILogger<WordExportService> logger)
{
    private static readonly ActivitySource ActivitySource = new("EtikettenGenerator.Web");

    private const int CellsPerPage = 40;

    /// <summary>
    /// Exports members to Word using the embedded label template.
    /// positionMemberPairs: 1-based position → member (ordered).
    /// </summary>
    public byte[] Export(IReadOnlyList<(int Position, Member Member)> positionMemberPairs,
        CancellationToken ct = default) =>
        ExportWithTemplate(positionMemberPairs, LoadTemplate(), ct);

    /// <summary>
    /// Exports members using a supplied template (for testing).
    /// </summary>
    internal byte[] ExportWithTemplate(IReadOnlyList<(int Position, Member Member)> positionMemberPairs,
        byte[] templateBytes, CancellationToken ct = default)
    {
        using var activity = ActivitySource.StartActivity("word.export");
        using var stream = new MemoryStream();
        stream.Write(templateBytes, 0, templateBytes.Length);
        stream.Position = 0;

        using (var doc = WordprocessingDocument.Open(stream, isEditable: true))
        {
            var body = doc.MainDocumentPart!.Document!.Body!;
            var templateTable = body.Elements<Table>().FirstOrDefault()
                ?? throw new InvalidOperationException("Vorlage nicht gefunden: Keine Tabelle im Dokument.");

            // Group into pages of 40
            var pages = positionMemberPairs
                .GroupBy(p => (p.Position - 1) / CellsPerPage)
                .OrderBy(g => g.Key)
                .ToList();

            var pageCount = pages.Count;

            // Fill the first page (template table)
            var firstPageLabels = GetLabelCells(templateTable);
            var firstPageFilled = pages[0].Select(p => p.Position).ToHashSet();
            FillPage(firstPageLabels, pages[0].ToList(), ct);
            ClearUnfilledLabels(firstPageLabels, firstPageFilled);

            // Append additional pages for overflow
            for (var i = 1; i < pages.Count; i++)
            {
                ct.ThrowIfCancellationRequested();

                // Add page break + cloned outer table
                body.AppendChild(new Paragraph(new Run(new Break { Type = BreakValues.Page })));
                var clonedTable = (Table)templateTable.CloneNode(deep: true);
                body.AppendChild(clonedTable);

                var pageLabelCells = GetLabelCells(clonedTable);
                // Normalize positions within this page to 1-40
                var pageEntries = pages[i]
                    .Select(p => (Position: ((p.Position - 1) % CellsPerPage) + 1, p.Member))
                    .ToList();
                var pageFilled = pageEntries.Select(p => p.Position).ToHashSet();
                FillPage(pageLabelCells, pageEntries, ct);
                ClearUnfilledLabels(pageLabelCells, pageFilled);
            }

            doc.Save();
        }

        var memberCount = positionMemberPairs.Count;
        var pageCount2 = (int)Math.Ceiling(memberCount / (double)CellsPerPage);
        activity?.SetTag("export.members", memberCount);
        activity?.SetTag("export.pages", pageCount2);
        logger.LogInformation("Word export: {Members} members, {Pages} pages", memberCount, pageCount2);

        return stream.ToArray();
    }

    private static void FillPage(List<TableCell> labelCells,
        IReadOnlyList<(int Position, Member Member)> entries,
        CancellationToken ct)
    {
        foreach (var (position, member) in entries)
        {
            ct.ThrowIfCancellationRequested();
            if (position < 1 || position > CellsPerPage || position - 1 >= labelCells.Count) continue;
            ReplaceMergeFields(labelCells[position - 1], BuildFieldValues(member));
        }
    }

    private static List<TableCell> GetLabelCells(Table outerTable) =>
        outerTable.Descendants<TableRow>()
            .SelectMany(row => row.Elements<TableCell>())
            .ToList();

    private static Dictionary<string, string> BuildFieldValues(Member member)
    {
        var (fs1, fs2, fs3) = ParseFahrerlaubnis(member.Fahrerlaubnis);
        return new Dictionary<string, string>
        {
            ["Nachname"] = member.Nachname,
            ["Vorname"]  = member.Vorname,
            ["Medizin"]  = MapMedQualifikation(member.MedQualifikation),
            ["Taktik"]   = MapDienststellung(member.Dienststellung),
            ["RD_Fortbildung"] = member.HatRettungsdienstfortbildung ? "RDF" : "",
            ["FS_1"]     = string.IsNullOrEmpty(fs1) ? "-" : fs1,
            ["FS_2"]     = string.IsNullOrEmpty(fs2) ? "-" : fs2,
            ["FS_3"]     = string.IsNullOrEmpty(fs3) ? "-" : fs3,
        };
    }

    internal static string MapMedQualifikation(string medQual) =>
        medQual.Trim() switch
        {
            var s when s.Equals("Rettungssanitäter/in",  StringComparison.OrdinalIgnoreCase) => "RS",
            var s when s.Equals("Sanitätshelfer/in",     StringComparison.OrdinalIgnoreCase) => "SH",
            var s when s.Equals("Notfallsanitäter/in",   StringComparison.OrdinalIgnoreCase) => "NFS",
            var s when s.Equals("Notarzt / Notärztin",   StringComparison.OrdinalIgnoreCase) => "NA",
            var s when s.Equals("Rettungshelfer/in",     StringComparison.OrdinalIgnoreCase) => "RH",
            var s when s.Equals("Erste-Hilfe",           StringComparison.OrdinalIgnoreCase) => "EH",
            var s when s.Equals("Rettungsassistent/in",  StringComparison.OrdinalIgnoreCase) => "RA",
            _ => "-",
        };

    internal static string MapDienststellung(string dienststellung) =>
        dienststellung.Trim() switch
        {
            var s when s.Equals("Helfer:in in Ausbildung", StringComparison.OrdinalIgnoreCase) => "HF",
            var s when s.Equals("ZF mit Stabsausbildung",  StringComparison.OrdinalIgnoreCase) => "GdSA",
            var s when s.Equals("Gruppenführer:in",        StringComparison.OrdinalIgnoreCase) => "GF",
            var s when s.Equals("Zugführer:in",            StringComparison.OrdinalIgnoreCase) => "ZF",
            var s when s.Equals("Verbandführer:in",        StringComparison.OrdinalIgnoreCase) => "VF",
            _ => "HF",
        };

    internal static (string Fs1, string Fs2, string Fs3) ParseFahrerlaubnis(string fahrerlaubnis)
    {
        if (string.IsNullOrWhiteSpace(fahrerlaubnis))
            return ("", "", "");

        var tokens = fahrerlaubnis
            .Split([' ', ',', ';'], StringSplitOptions.RemoveEmptyEntries)
            .Select(t => t.Trim())
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        // FS_1: highest PKW/LKW class, E-variants excluded
        var fs1Hierarchy = new[] { "C", "C1", "B", "B96" };
        var eClasses = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "CE", "C1E", "BE", "B+E" };
        var fs1 = fs1Hierarchy.FirstOrDefault(c => tokens.Contains(c) && !eClasses.Contains(c)) ?? "";

        // FS_2: "E" if any E-class present
        var fs2 = tokens.Any(t => eClasses.Contains(t)) ? "E" : "";

        // FS_3: highest motorcycle class (AM excluded — not relevant)
        var fs3Hierarchy = new[] { "A", "A2", "A1" };
        var fs3 = fs3Hierarchy.FirstOrDefault(c => tokens.Contains(c)) ?? "";

        return (fs1, fs2, fs3);
    }

    private static void ReplaceMergeFields(OpenXmlElement container,
        Dictionary<string, string> fieldValues)
    {
        foreach (var para in container.Descendants<Paragraph>())
        {
            string? fieldName = null;
            var inDisplaySection = false;

            foreach (var run in para.Elements<Run>())
            {
                var fldChar  = run.GetFirstChild<FieldChar>();
                var instrTxt = run.GetFirstChild<FieldCode>();

                if (fldChar?.FieldCharType?.Value == FieldCharValues.Begin)
                {
                    fieldName = null;
                    inDisplaySection = false;
                }
                else if (instrTxt != null)
                {
                    var instr = instrTxt.Text.Trim();
                    if (instr.StartsWith("MERGEFIELD ", StringComparison.OrdinalIgnoreCase))
                        fieldName = instr["MERGEFIELD ".Length..].Trim();
                }
                else if (fldChar?.FieldCharType?.Value == FieldCharValues.Separate)
                {
                    inDisplaySection = true;
                }
                else if (fldChar?.FieldCharType?.Value == FieldCharValues.End)
                {
                    inDisplaySection = false;
                    fieldName = null;
                }
                else if (inDisplaySection && fieldName != null)
                {
                    if (fieldValues.TryGetValue(fieldName, out var value))
                    {
                        var textEl = run.GetFirstChild<Text>();
                        if (textEl != null)
                        {
                            textEl.Text = value;
                            textEl.Space = SpaceProcessingModeValues.Preserve;
                        }
                    }
                }
            }
        }
    }

    private static void ClearMergeFields(OpenXmlElement container)
    {
        foreach (var para in container.Descendants<Paragraph>())
        {
            var inDisplaySection = false;
            var hasFieldName = false;

            foreach (var run in para.Elements<Run>())
            {
                var fldChar  = run.GetFirstChild<FieldChar>();
                var instrTxt = run.GetFirstChild<FieldCode>();

                if (fldChar?.FieldCharType?.Value == FieldCharValues.Begin)
                {
                    hasFieldName = false;
                    inDisplaySection = false;
                }
                else if (instrTxt != null)
                {
                    hasFieldName = instrTxt.Text.Trim()
                        .StartsWith("MERGEFIELD ", StringComparison.OrdinalIgnoreCase);
                }
                else if (fldChar?.FieldCharType?.Value == FieldCharValues.Separate)
                {
                    inDisplaySection = true;
                }
                else if (fldChar?.FieldCharType?.Value == FieldCharValues.End)
                {
                    inDisplaySection = false;
                    hasFieldName = false;
                }
                else if (inDisplaySection && hasFieldName)
                {
                    var textEl = run.GetFirstChild<Text>();
                    if (textEl != null)
                    {
                        textEl.Text = "";
                        textEl.Space = SpaceProcessingModeValues.Preserve;
                    }
                }
            }
        }
    }

    private static void ClearUnfilledLabels(List<TableCell> labelCells, HashSet<int> filledPositions)
    {
        for (var i = 0; i < labelCells.Count; i++)
        {
            if (!filledPositions.Contains(i + 1))
                ClearMergeFields(labelCells[i]);
        }
    }

    private static byte[] LoadTemplate()
    {
        var assembly = typeof(WordExportService).Assembly;
        const string resourceName = "EtikettenGenerator.Web.Templates.EtikettenVorlage.docx";

        using var stream = assembly.GetManifestResourceStream(resourceName)
            ?? throw new InvalidOperationException($"Vorlage nicht gefunden: {resourceName}");

        using var ms = new MemoryStream();
        stream.CopyTo(ms);
        return ms.ToArray();
    }
}
