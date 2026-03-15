# Spec: Blazor Etiketten-Generator mit .NET Aspire

## Projektübersicht

Eine .NET 9-basierte Blazor-Web-App (Blazor Server / Interactive Server rendering), eingebettet in eine **.NET Aspire 9.x**-Solution. Die App liest eine CSV-Datei ein, zeigt Mitglieder selektierbar an und befüllt eine Word-Vorlage mit Etiketten-Daten auf einem 4×10-Raster (40 Etiketten pro Seite).

---

## Technologie-Stack

| Bereich | Technologie |
|---|---|
| Plattform | **.NET 9** |
| Orchestrierung | **.NET Aspire 9.x** (aktuell 9.5.x) |
| Frontend | **Blazor Web App** (Interactive Server, .NET 9) |
| CSV-Parsing | **CsvHelper** (NuGet) |
| Word-Export | **DocumentFormat.OpenXml** (Open XML SDK, NuGet) |
| Observability | OpenTelemetry via Aspire ServiceDefaults (automatisch) |
| Health Checks | Aspire ServiceDefaults (automatisch) |
| Resilience | Microsoft.Extensions.Http.Resilience via ServiceDefaults |
| UI-Komponenten | **MudBlazor** oder **Radzen Blazor** |
| Aspire CLI | `aspire new` / `aspire run` / `aspire publish` |

---

## Solution-Struktur (.NET Aspire Standard-Layout)

```
EtikettenGenerator.sln
│
├── EtikettenGenerator.AppHost/          ← Aspire AppHost (Orchestrator)
│   ├── EtikettenGenerator.AppHost.csproj
│   └── Program.cs                       ← Topology-Definition
│
├── EtikettenGenerator.ServiceDefaults/  ← Aspire ServiceDefaults (shared)
│   ├── EtikettenGenerator.ServiceDefaults.csproj
│   └── Extensions.cs                    ← OTel, Health, Resilience
│
└── EtikettenGenerator.Web/              ← Blazor Web App (die eigentliche App)
    ├── EtikettenGenerator.Web.csproj
    ├── Program.cs
    ├── Pages/
    │   └── Index.razor
    ├── Components/
    │   ├── CsvUpload.razor
    │   ├── MemberTable.razor
    │   └── PositionPickerDialog.razor
    ├── Models/
    │   └── Member.cs
    ├── Services/
    │   ├── CsvImportService.cs
    │   └── WordExportService.cs
    └── Templates/
        └── EtikettenVorlage.docx
```

---

## .NET Aspire Konfiguration

### `AppHost/Program.cs`

```csharp
var builder = DistributedApplication.CreateBuilder(args);

var web = builder.AddProject<Projects.EtikettenGenerator_Web>("web")
    .WithExternalHttpEndpoints();

await builder.Build().RunAsync();
```

> **Hinweis:** Da diese App keine externen Dienste (Datenbank, Redis, API) benötigt,
> ist das AppHost-Modell bewusst schlank. Der Nutzen liegt in:
> - Einheitlichem Start über `aspire run` / F5
> - Aspire Dashboard mit Logs, Traces, Metriken, Health-Status
> - Vorbereitung für spätere Erweiterungen (z.B. API-Service auslagern)

### `ServiceDefaults/Extensions.cs`

Standard-Aspire-Vorlage ohne Anpassung:
- `AddServiceDefaults()` auf dem `IHostApplicationBuilder`
- Aktiviert: OpenTelemetry Logging, Metrics, Tracing, Health Checks, HTTP Resilience, Service Discovery

### `Web/Program.cs`

```csharp
var builder = WebApplication.CreateBuilder(args);

// Aspire ServiceDefaults einbinden
builder.AddServiceDefaults();

// Blazor
builder.Services.AddRazorComponents()
    .AddInteractiveServerComponents();

// App-Services
builder.Services.AddScoped<CsvImportService>();
builder.Services.AddScoped<WordExportService>();

var app = builder.Build();

app.MapDefaultEndpoints(); // Aspire Health-Check-Endpoints
app.UseAntiforgery();
app.MapRazorComponents<App>()
    .AddInteractiveServerRenderMode();

app.Run();
```

---

## Datenmodell

### `Member.cs`

```csharp
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
```

> Verwendung von `record` + `init`-Properties (C# 9+) für Immutability der Importdaten.

---

## CSV-Import (`CsvImportService.cs`)

### Spalten-Mapping (fest definiert)

| CSV-Spaltenname      | Modell-Feld                 |
|----------------------|-----------------------------|
| `Nachname`           | `Member.Nachname`           |
| `Vorname`            | `Member.Vorname`            |
| `med. Qualifikation` | `Member.MedQualifikation`   |
| `Dienststellung`     | `Member.Dienststellung`     |
| `Fahrerlaubnis`      | `Member.Fahrerlaubnis`      |
| `Ausbildungen`       | → `HatRettungsdienstfortbildung` |

### Implementierungsanforderungen

- **Encoding**: UTF-8 mit BOM-Unterstützung; Fallback auf Latin-1 (`ISO-8859-1`)
- **Trennzeichen**: Semikolon (`;`) als Standard; automatische Erkennung von Komma (`,`) als Fallback
- **Fehlende Spalten**: `CsvImportException` mit Liste fehlender Spaltennamen werfen
- **Leere Zeilen**: überspringen
- **Ausbildungen-Prüfung**: `StringComparison.OrdinalIgnoreCase` auf Teilstring `"Rettungsdienstfortbildung"`
- Rückgabe: `IReadOnlyList<Member>`

---

## UI – Hauptseite (`Index.razor`)

### Layout

```
┌─────────────────────────────────────────────────────────┐
│  📂 CSV-Datei hochladen: [Datei auswählen]              │
│  ⚠️ [Fehlermeldungen / Warnungen]                       │
├─────────────────────────────────────────────────────────┤
│  [☑ Alle auswählen]              [🏷️ Etiketten erstellen] │
├──┬────────────┬──────────┬────────────┬─────────────────┤
│☑ │ Nachname   │ Vorname  │ med. Qual. │ Dienststell. │ … │
├──┼────────────┼──────────┼────────────┼─────────────────┤
│☑ │ Mustermann │ Max      │ Notarzt    │ …            │ … │
│☐ │ …          │ …        │ …          │ …            │ … │
└──┴────────────┴──────────┴────────────┴─────────────────┘
```

### Tabellenspalten

1. Checkbox (Selektion)
2. Nachname
3. Vorname
4. med. Qualifikation
5. Dienststellung
6. Fahrerlaubnis
7. Rettungsdienstfortbildung (`✓` / `✗`)

### Besondere UI-Anforderungen

- Datei-Upload via `<InputFile>` (Blazor nativ, kein JS-Interop nötig)
- Fortschrittsanzeige / Spinner während CSV-Parsing und Word-Export (`isLoading`-State)
- Fehlermeldungen in einer `MudAlert`-Komponente oberhalb der Tabelle
- Tabelle scrollbar bei vielen Einträgen
- „Alle auswählen" Checkbox: tri-state (alle / keine / gemischt)

---

## Button „Etiketten erstellen" – Logik

```
Klick auf "Etiketten erstellen"
│
├── Keine Mitglieder geladen?
│   → Alert: "Bitte zuerst eine CSV-Datei laden."
│
├── Keine Mitglieder selektiert?
│   → Alert: "Bitte mindestens ein Mitglied auswählen."
│
├── Alle vorhandenen Mitglieder selektiert?
│   → Word-Export direkt starten
│       → Positionen 1..N sequenziell füllen
│       → Bei N > 40: automatisch neue Seite(n) anhängen
│
└── Teilmenge selektiert?
    → PositionPickerDialog öffnen
        → Benutzer wählt Startpositionen im 4×10-Raster
        → [Bestätigen] → Word-Export mit gewählter Positionszuordnung
        → [Abbrechen] → Dialog schließen, kein Export
```

---

## Dialog: Rasterposition-Auswahl (`PositionPickerDialog.razor`)

### Verhalten

- Zeigt ein **4 Spalten × 10 Zeilen = 40 Felder**-Raster
- Nummerierung: 1–40, links-nach-rechts, oben-nach-unten
- Jedes Feld: togglebar (Klick = ausgewählt/deselektiert)
- Visuelles Feedback: ausgewählte Felder grün hervorgehoben
- **Anzahl wählbarer Positionen = Anzahl selektierter Mitglieder** (harte Obergrenze)
- Wird die Obergrenze erreicht: weitere Felder deaktiviert (grau, nicht anklickbar)
- **[Bestätigen]** nur aktiv wenn: `ausgewähltePositionen.Count == selektierteMitglieder.Count`
- Zeigt Statuszeile: „_X von Y Positionen gewählt_"

### Raster-Nummerierung

```
 1  |  2  |  3  |  4
 5  |  6  |  7  |  8
 9  | 10  | 11  | 12
...
37  | 38  | 39  | 40
```

---

## Word-Export (`WordExportService.cs`)

### Vorlage (`EtikettenVorlage.docx`)

- A4-Seite (Hochformat empfohlen)
- Enthält eine Word-Tabelle mit genau **10 Zeilen × 4 Spalten = 40 Zellen**
- Jede Zelle enthält Text-Marker (einfach zu pflegen in Word):

```
{{NACHNAME}}
{{VORNAME}}
{{MED_QUAL}}
{{DIENSTSTELLUNG}}
{{FAHRERLAUBNIS}}
{{RD_FORTBILDUNG}}
```

### Export-Logik (Open XML SDK)

```csharp
// Pseudocode
using var stream = new MemoryStream(templateBytes);
using var doc = WordprocessingDocument.Open(stream, isEditable: true);

var table = doc.MainDocumentPart!.Document.Body!
    .Elements<Table>().First();

var cells = table.Descendants<TableCell>().ToList(); // 40 Zellen

for (int i = 0; i < positionMemberPairs.Count; i++)
{
    var (position, member) = positionMemberPairs[i];
    var cell = cells[position - 1]; // Position ist 1-basiert
    ReplacePlaceholders(cell, member);
}

// Leere Positionen: Marker entfernen (Zelle bleibt leer)
foreach (var emptyCell in cells.Where(c => NochPlatzhalterVorhanden(c)))
    ClearPlaceholders(emptyCell);

doc.Save();
return stream.ToArray();
```

### Platzhalter-Ersetzung

- Alle `<w:t>`-Textruns in der Zelle durchsuchen
- **Run-Merge**: Platzhalter können auf mehrere Runs verteilt sein → Runs in der Zelle vorher zusammenführen (`MergeAdjacentRuns`)
- Exaktes String-Matching auf `{{...}}`-Marker
- `HatRettungsdienstfortbildung`: `true` → `"Ja"` / `false` → `""` (leer, keine negative Aussage)

### Überlauf auf weitere Seiten (>40 Mitglieder)

Bei Vollexport mit mehr als 40 Mitgliedern:
1. Ersten 40 auf Seite 1 eintragen
2. Für jede weitere 40er-Gruppe: eine Kopie der Vorlage-Seite (Tabelle + Seitenumbruch) ans Dokument anhängen
3. Die kopierten Tabellen befüllen

### Download (Blazor Server)

```csharp
// In der Razor-Komponente via IJSRuntime
var bytes = await WordExportService.ExportAsync(members, positions);
var fileName = $"Etiketten_{DateTime.Now:yyyy-MM-dd}.docx";
await JS.InvokeVoidAsync("downloadFile", fileName,
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    Convert.ToBase64String(bytes));
```

JS-Hilfsfunktion in `wwwroot/app.js`:
```javascript
window.downloadFile = (filename, contentType, base64) => {
    const a = document.createElement('a');
    a.href = `data:${contentType};base64,${base64}`;
    a.download = filename;
    a.click();
};
```

---

## Observability via Aspire Dashboard

Dank `AddServiceDefaults()` stehen im **Aspire Dashboard** automatisch bereit:

| Feature | Details |
|---|---|
| **Structured Logging** | Alle `ILogger<T>`-Logs aus der Web-App sichtbar |
| **Distributed Tracing** | Request-Traces inkl. CSV-Import und Word-Export |
| **Metrics** | ASP.NET Core, HTTP Client, Runtime-Metriken |
| **Health Checks** | `/health` und `/alive` Endpoints |
| **Resource View** | Live-Status der Web-App im Dashboard |

Eigene Telemetrie-Instrumentierung empfohlen für:
```csharp
// In CsvImportService
using var activity = ActivitySource.StartActivity("csv.import");
activity?.SetTag("csv.rows", members.Count);

// In WordExportService
using var activity = ActivitySource.StartActivity("word.export");
activity?.SetTag("export.members", members.Count);
activity?.SetTag("export.pages", pageCount);
```

---

## NuGet-Pakete

### `EtikettenGenerator.AppHost.csproj`
```xml
<PackageReference Include="Aspire.Hosting.AppHost" Version="9.*" />
```

### `EtikettenGenerator.ServiceDefaults.csproj`
```xml
<PackageReference Include="Microsoft.Extensions.Http.Resilience" Version="9.*" />
<PackageReference Include="Microsoft.Extensions.ServiceDiscovery" Version="9.*" />
<PackageReference Include="OpenTelemetry.Exporter.OpenTelemetryProtocol" Version="1.*" />
<PackageReference Include="OpenTelemetry.Extensions.Hosting" Version="1.*" />
<PackageReference Include="OpenTelemetry.Instrumentation.AspNetCore" Version="1.*" />
<PackageReference Include="OpenTelemetry.Instrumentation.Http" Version="1.*" />
<PackageReference Include="OpenTelemetry.Instrumentation.Runtime" Version="1.*" />
```

### `EtikettenGenerator.Web.csproj`
```xml
<PackageReference Include="CsvHelper" Version="33.*" />
<PackageReference Include="DocumentFormat.OpenXml" Version="3.*" />
<PackageReference Include="MudBlazor" Version="7.*" />
<!-- ServiceDefaults als Projektreferenz -->
<ProjectReference Include="..\EtikettenGenerator.ServiceDefaults\EtikettenGenerator.ServiceDefaults.csproj" />
```

---

## Fehlerbehandlung

| Fehlerfall | Verhalten |
|---|---|
| CSV-Spalte fehlt | `MudAlert` Severity.Error mit fehlenden Spaltennamen |
| CSV leer | Alert: „Keine Datensätze gefunden." |
| Keine Selektion beim Export | Alert: „Bitte mindestens ein Mitglied auswählen." |
| Word-Vorlage nicht gefunden | Alert: „Vorlage nicht gefunden." + Logging via `ILogger` |
| Allgemeiner Fehler | Alert + `ILogger.LogError` → sichtbar im Aspire Dashboard |

---

## Nicht-funktionale Anforderungen

- **Keine persistente Datenspeicherung**: Alle Daten nur im Scoped-DI-Service (Session-Scoped bei Blazor Server)
- **Lokalisierung**: Deutsch (DE), kein i18n-Framework
- **Aspire CLI Kompatibilität**: `aspire run` startet die gesamte Solution
- **Ziel-Framework**: `net9.0`; bereit für spätere Migration auf `net10.0` (LTS)
- **C# Version**: C# 13 (mit .NET 9)
- **Nullable Reference Types**: `<Nullable>enable</Nullable>` in allen Projekten
- **Barrierefreiheit**: `aria-label` auf interaktiven Elementen

---

## Offene Entscheidungen

| # | Frage | Empfehlung |
|---|---|---|
| 1 | Blazor Server vs. WASM | **Blazor Server** (einfacheres Datei-Streaming, direkte Aspire-Integration) |
| 2 | UI-Bibliothek | **MudBlazor** (aktiv gepflegt, gute Blazor-Server-Kompatibilität) |
| 3 | Platzhalter in Vorlage | **`{{TEXT}}`-Marker** (einfach im Word-Editor pflegbar) |
| 4 | Word-Bibliothek | **Open XML SDK** (kostenlos, offiziell von Microsoft) |
| 5 | Überlauf >40 Mitglieder | Neue Seite automatisch anhängen (Kopie der Vorlage-Tabelle) |
| 6 | Aspire Deployment-Ziel | Docker Compose (`aspire publish`) oder Azure Container Apps |

---

## Setup & Erste Schritte (Entwickler)

```bash
# Aspire CLI installieren (falls nicht vorhanden)
curl -sSL https://aspire.dev/install.sh | bash        # Linux/macOS
iex "& { $(irm https://aspire.dev/install.ps1) }"    # Windows PowerShell

# Solution erstellen (Aspire Starter-Template)
aspire new aspire-starter -o EtikettenGenerator
cd EtikettenGenerator

# Abhängige NuGet-Pakete hinzufügen (im Web-Projekt)
dotnet add EtikettenGenerator.Web package CsvHelper
dotnet add EtikettenGenerator.Web package DocumentFormat.OpenXml
dotnet add EtikettenGenerator.Web package MudBlazor

# Solution starten
aspire run
# → Aspire Dashboard öffnet sich automatisch im Browser
# → Blazor App unter https://localhost:{port} erreichbar
```

---

## Akzeptanzkriterien

- [ ] `aspire run` startet die gesamte Solution ohne weitere Konfiguration
- [ ] Aspire Dashboard zeigt die Web-App mit Health-Status „Running"
- [ ] Logs aus CSV-Import und Word-Export sind im Aspire Dashboard sichtbar
- [ ] CSV-Datei mit den definierten Spalten wird korrekt eingelesen
- [ ] Spalte „Ausbildungen" wird korrekt auf „Rettungsdienstfortbildung" geprüft
- [ ] Alle Mitglieder werden in der Tabelle angezeigt und sind selektierbar
- [ ] „Alle auswählen" funktioniert korrekt (inkl. tri-state)
- [ ] Bei Vollselektion wird direkt exportiert, kein Dialog
- [ ] Bei Teilselektion erscheint der Positionsdialog mit 4×10-Raster
- [ ] [Bestätigen] im Dialog ist nur aktiv wenn alle Positionen zugewiesen sind
- [ ] Das erzeugte Word-Dokument enthält die richtigen Daten an den richtigen Positionen
- [ ] Leere Positionen enthalten keinen Platzhalter-Text
- [ ] Bei >40 Mitgliedern werden weitere Seiten automatisch angehängt
- [ ] Download des fertigen `.docx` funktioniert im Browser