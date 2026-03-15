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
            var firstPageCells = GetCells(templateTable);
            FillPage(firstPageCells, pages[0].ToList(), ct);

            // Append additional pages for overflow
            for (var i = 1; i < pages.Count; i++)
            {
                ct.ThrowIfCancellationRequested();

                // Add page break + cloned table
                body.AppendChild(new Paragraph(new Run(new Break { Type = BreakValues.Page })));
                var clonedTable = (Table)templateTable.CloneNode(deep: true);
                body.AppendChild(clonedTable);

                var pageCells = GetCells(clonedTable);
                // Normalize positions within this page to 1-40
                var pageEntries = pages[i]
                    .Select(p => (Position: ((p.Position - 1) % CellsPerPage) + 1, p.Member))
                    .ToList();
                FillPage(pageCells, pageEntries, ct);
            }

            // Clear any remaining unfilled placeholders in all tables
            foreach (var table in body.Elements<Table>())
            {
                foreach (var cell in table.Descendants<TableCell>())
                {
                    ClearPlaceholders(cell);
                }
            }

            doc.Save();
        }

        var result = stream.ToArray();
        var memberCount = positionMemberPairs.Count;
        var pageCount2 = (int)Math.Ceiling(memberCount / (double)CellsPerPage);
        activity?.SetTag("export.members", memberCount);
        activity?.SetTag("export.pages", pageCount2);
        logger.LogInformation("Word export: {Members} members, {Pages} pages", memberCount, pageCount2);

        return result;
    }

    private static void FillPage(List<TableCell> cells,
        IReadOnlyList<(int Position, Member Member)> entries,
        CancellationToken ct)
    {
        foreach (var (position, member) in entries)
        {
            ct.ThrowIfCancellationRequested();
            if (position < 1 || position > CellsPerPage) continue;
            var cell = cells[position - 1];
            MergeAdjacentRuns(cell);
            ReplacePlaceholders(cell, member);
        }
    }

    private static List<TableCell> GetCells(Table table) =>
        table.Descendants<TableCell>().ToList();

    private static void ReplacePlaceholders(TableCell cell, Member member)
    {
        var replacements = new Dictionary<string, string>
        {
            ["{{NACHNAME}}"] = member.Nachname,
            ["{{VORNAME}}"] = member.Vorname,
            ["{{MED_QUAL}}"] = member.MedQualifikation,
            ["{{DIENSTSTELLUNG}}"] = member.Dienststellung,
            ["{{FAHRERLAUBNIS}}"] = member.Fahrerlaubnis,
            ["{{RD_FORTBILDUNG}}"] = member.HatRettungsdienstfortbildung ? "Ja" : "",
        };

        foreach (var textElement in cell.Descendants<Text>())
        {
            foreach (var (placeholder, value) in replacements)
            {
                if (textElement.Text.Contains(placeholder))
                    textElement.Text = textElement.Text.Replace(placeholder, value);
            }
        }
    }

    private static void ClearPlaceholders(TableCell cell)
    {
        foreach (var textElement in cell.Descendants<Text>())
        {
            if (textElement.Text.StartsWith("{{") && textElement.Text.EndsWith("}}"))
                textElement.Text = "";
        }
    }

    private static void MergeAdjacentRuns(TableCell cell)
    {
        foreach (var paragraph in cell.Descendants<Paragraph>())
        {
            var runs = paragraph.Elements<Run>().ToList();
            if (runs.Count <= 1) continue;

            // Merge all runs into the first one
            var first = runs[0];
            var firstText = first.GetFirstChild<Text>();
            if (firstText is null)
            {
                firstText = new Text();
                first.AppendChild(firstText);
            }

            var combined = string.Concat(runs.Select(r => r.GetFirstChild<Text>()?.Text ?? ""));
            firstText.Text = combined;
            firstText.Space = SpaceProcessingModeValues.Preserve;

            for (var i = 1; i < runs.Count; i++)
                runs[i].Remove();
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
