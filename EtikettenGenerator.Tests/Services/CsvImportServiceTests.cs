using System.Text;
using AwesomeAssertions;
using EtikettenGenerator.Web.Services;
using Microsoft.Extensions.Logging.Abstractions;

namespace EtikettenGenerator.Tests.Services;

public sealed class CsvImportServiceTests
{
    private static CsvImportService CreateService() =>
        new(NullLogger<CsvImportService>.Instance);

    private static Stream ToCsvStream(string csv, Encoding? encoding = null) =>
        new MemoryStream((encoding ?? Encoding.UTF8).GetBytes(csv));

    private const string ValidCsvSemicolon =
        "Nachname;Vorname;med. Qualifikation;Dienststellung;Fahrerlaubnis;Ausbildungen\n" +
        "Mustermann;Max;Notarzt;Staffelführer;B;Rettungsdienstfortbildung 2024\n" +
        "Schmidt;Anna;Rettungssanitäter;Gruppenführer;C;Brandschutz\n";

    private const string ValidCsvComma =
        "Nachname,Vorname,med. Qualifikation,Dienststellung,Fahrerlaubnis,Ausbildungen\n" +
        "Mustermann,Max,Notarzt,Staffelführer,B,Rettungsdienstfortbildung\n";

    [Fact]
    public async Task ImportAsync_ValidSemicolonCsv_ReturnsParsedMembers()
    {
        var service = CreateService();
        using var stream = ToCsvStream(ValidCsvSemicolon);

        var members = await service.ImportAsync(stream);

        members.Should().HaveCount(2);
        members[0].Nachname.Should().Be("Mustermann");
        members[0].Vorname.Should().Be("Max");
        members[0].HatRettungsdienstfortbildung.Should().BeTrue();
        members[1].HatRettungsdienstfortbildung.Should().BeFalse();
    }

    [Fact]
    public async Task ImportAsync_CommaDelimiter_ParsesCorrectly()
    {
        var service = CreateService();
        using var stream = ToCsvStream(ValidCsvComma);

        var members = await service.ImportAsync(stream);

        members.Should().HaveCount(1);
        members[0].Nachname.Should().Be("Mustermann");
    }

    [Fact]
    public async Task ImportAsync_Utf8BomEncoding_ParsesCorrectly()
    {
        var service = CreateService();
        var bom = new byte[] { 0xEF, 0xBB, 0xBF };
        var csvBytes = bom.Concat(Encoding.UTF8.GetBytes(ValidCsvSemicolon)).ToArray();
        using var stream = new MemoryStream(csvBytes);

        var members = await service.ImportAsync(stream);

        members.Should().NotBeEmpty();
    }

    [Fact]
    public async Task ImportAsync_MissingRequiredColumn_ThrowsCsvImportException()
    {
        var service = CreateService();
        const string csv = "Nachname;Vorname;Dienststellung\nMustermann;Max;Staffelführer\n";
        using var stream = ToCsvStream(csv);

        var act = () => service.ImportAsync(stream);

        await act.Should().ThrowAsync<CsvImportException>()
            .WithMessage("*Fehlende Spalten*");
    }

    [Fact]
    public async Task ImportAsync_MissingColumn_ExceptionListsMissingColumnNames()
    {
        var service = CreateService();
        const string csv = "Nachname;Vorname\nMustermann;Max\n";
        using var stream = ToCsvStream(csv);

        var exception = await Assert.ThrowsAsync<CsvImportException>(() => service.ImportAsync(stream));

        exception.MissingColumns.Should().Contain("med. Qualifikation");
        exception.MissingColumns.Should().Contain("Fahrerlaubnis");
        exception.MissingColumns.Should().Contain("Ausbildungen");
    }

    [Fact]
    public async Task ImportAsync_EmptyFile_ThrowsCsvImportException()
    {
        var service = CreateService();
        const string csv = "Nachname;Vorname;med. Qualifikation;Dienststellung;Fahrerlaubnis;Ausbildungen\n";
        using var stream = ToCsvStream(csv);

        var act = () => service.ImportAsync(stream);

        await act.Should().ThrowAsync<CsvImportException>()
            .WithMessage("*Keine Datensätze*");
    }

    [Fact]
    public async Task ImportAsync_RettungsdienstfortbildungCaseInsensitive_DetectedCorrectly()
    {
        var service = CreateService();
        const string csv =
            "Nachname;Vorname;med. Qualifikation;Dienststellung;Fahrerlaubnis;Ausbildungen\n" +
            "Test;User;RS;Staffelführer;B;RETTUNGSDIENSTFORTBILDUNG 2023\n";
        using var stream = ToCsvStream(csv);

        var members = await service.ImportAsync(stream);

        members[0].HatRettungsdienstfortbildung.Should().BeTrue();
    }

    [Fact]
    public async Task ImportAsync_DefaultSelectedState_IsTrue()
    {
        var service = CreateService();
        using var stream = ToCsvStream(ValidCsvSemicolon);

        var members = await service.ImportAsync(stream);

        members.Should().AllSatisfy(m => m.IsSelected.Should().BeTrue());
    }
}
