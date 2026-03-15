namespace EtikettenGenerator.Web.Models;

public sealed record Member
{
    public string Nachname { get; init; } = "";
    public string Vorname { get; init; } = "";
    public string MedQualifikation { get; init; } = "";
    public string Dienststellung { get; init; } = "";
    public string Fahrerlaubnis { get; init; } = "";
    public bool HatRettungsdienstfortbildung { get; init; }
    public bool IsSelected { get; set; } = true;
}
