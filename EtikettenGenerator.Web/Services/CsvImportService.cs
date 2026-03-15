using System.Diagnostics;
using System.Text;
using CsvHelper;
using CsvHelper.Configuration;
using EtikettenGenerator.Web.Models;
using Microsoft.Extensions.Logging;

namespace EtikettenGenerator.Web.Services;

public sealed class CsvImportService(ILogger<CsvImportService> logger)
{
    private static readonly ActivitySource ActivitySource = new("EtikettenGenerator.Web");

    private static readonly string[] RequiredColumns =
    [
        "Nachname", "Vorname", "med. Qualifikation", "Dienststellung", "Fahrerlaubnis", "Ausbildungen"
    ];

    public async Task<IReadOnlyList<Member>> ImportAsync(Stream stream, CancellationToken ct = default)
    {
        using var activity = ActivitySource.StartActivity("csv.import");

        var bytes = await ReadAllBytesAsync(stream, ct);

        var (encoding, delimiter) = DetectEncodingAndDelimiter(bytes);
        logger.LogInformation("CSV import: detected encoding={Encoding}, delimiter='{Delimiter}'",
            encoding.WebName, delimiter);

        using var reader = new StreamReader(new MemoryStream(bytes), encoding);
        var config = new CsvConfiguration(System.Globalization.CultureInfo.InvariantCulture)
        {
            Delimiter = delimiter,
            HasHeaderRecord = true,
            MissingFieldFound = null,
            BadDataFound = null,
            TrimOptions = TrimOptions.Trim,
        };

        using var csv = new CsvReader(reader, config);
        await csv.ReadAsync();
        csv.ReadHeader();

        ValidateHeaders(csv.HeaderRecord ?? []);

        var members = new List<Member>();
        while (await csv.ReadAsync())
        {
            ct.ThrowIfCancellationRequested();

            var ausbildungen = csv.GetField("Ausbildungen") ?? "";
            members.Add(new Member
            {
                Nachname = csv.GetField("Nachname") ?? "",
                Vorname = csv.GetField("Vorname") ?? "",
                MedQualifikation = csv.GetField("med. Qualifikation") ?? "",
                Dienststellung = csv.GetField("Dienststellung") ?? "",
                Fahrerlaubnis = csv.GetField("Fahrerlaubnis") ?? "",
                HatRettungsdienstfortbildung = ausbildungen.Contains(
                    "Rettungsdienstfortbildung", StringComparison.OrdinalIgnoreCase),
            });
        }

        if (members.Count == 0)
            throw new CsvImportException("Keine Datensätze gefunden.");

        activity?.SetTag("csv.rows", members.Count);
        logger.LogInformation("CSV import: {Count} members imported", members.Count);

        return members;
    }

    private static async Task<byte[]> ReadAllBytesAsync(Stream stream, CancellationToken ct)
    {
        using var ms = new MemoryStream();
        await stream.CopyToAsync(ms, ct);
        return ms.ToArray();
    }

    private static (Encoding encoding, string delimiter) DetectEncodingAndDelimiter(byte[] bytes)
    {
        Encoding encoding;

        // Check for UTF-8 BOM
        if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF)
        {
            encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: true);
        }
        else
        {
            // Try UTF-8, fall back to Latin-1
            try
            {
                var utf8 = Encoding.UTF8;
                utf8.GetString(bytes); // will throw if invalid UTF-8 sequence
                encoding = utf8;
            }
            catch
            {
                encoding = Encoding.Latin1;
            }
        }

        var sample = encoding.GetString(bytes[..Math.Min(bytes.Length, 1024)]);
        var semicolonCount = sample.Count(c => c == ';');
        var commaCount = sample.Count(c => c == ',');
        var delimiter = semicolonCount >= commaCount ? ";" : ",";

        return (encoding, delimiter);
    }

    private static void ValidateHeaders(string[] headers)
    {
        var missing = RequiredColumns.Except(headers, StringComparer.OrdinalIgnoreCase).ToList();
        if (missing.Count > 0)
            throw new CsvImportException($"Fehlende Spalten: {string.Join(", ", missing)}", missing);
    }
}

public sealed class CsvImportException : Exception
{
    public IReadOnlyList<string> MissingColumns { get; }

    public CsvImportException(string message, IReadOnlyList<string>? missingColumns = null)
        : base(message)
    {
        MissingColumns = missingColumns ?? [];
    }
}
