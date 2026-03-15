# CLAUDE.md — EtikettenGenerator

## Project Overview

A **.NET 10 Blazor Server** application embedded in a **.NET Aspire 9.x** solution.
Reads member data from CSV, displays them in a selectable table, and exports a populated Word label template (4×10 grid, 40 labels per page) as a `.docx` download.

**Domain:** Emergency dispatch board labels (Einsatzführungsboard-Etiketten)
**Language:** German UI and domain terms; code in English

---

## Solution Structure

```
EtikettenGenerator.sln
├── EtikettenGenerator.AppHost/           ← Aspire orchestrator
│   └── Program.cs
├── EtikettenGenerator.ServiceDefaults/   ← Shared OTel / health / resilience
│   └── Extensions.cs
└── EtikettenGenerator.Web/               ← Blazor Server app
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

## Technology Stack

| Concern         | Technology                                       |
|-----------------|--------------------------------------------------|
| Platform        | **.NET 10** (`net10.0`)                          |
| Orchestration   | **.NET Aspire 9.x** (≥ 9.5.x)                   |
| Frontend        | **Blazor Web App** — Interactive Server render   |
| UI Library      | **MudBlazor 7.x** (custom CI theme, see below)  |
| CSV Parsing     | **CsvHelper 33.x**                               |
| Word Export     | **DocumentFormat.OpenXml 3.x** (Open XML SDK)   |
| Observability   | OpenTelemetry via Aspire ServiceDefaults         |
| Testing         | **xUnit** + **AwesomeAssertions** (FA fork)      |

---

## NuGet Packages

### AppHost
```xml
<PackageReference Include="Aspire.Hosting.AppHost" Version="9.*" />
```

### ServiceDefaults
```xml
<PackageReference Include="Microsoft.Extensions.Http.Resilience" Version="9.*" />
<PackageReference Include="Microsoft.Extensions.ServiceDiscovery" Version="9.*" />
<PackageReference Include="OpenTelemetry.Exporter.OpenTelemetryProtocol" Version="1.*" />
<PackageReference Include="OpenTelemetry.Extensions.Hosting" Version="1.*" />
<PackageReference Include="OpenTelemetry.Instrumentation.AspNetCore" Version="1.*" />
<PackageReference Include="OpenTelemetry.Instrumentation.Http" Version="1.*" />
<PackageReference Include="OpenTelemetry.Instrumentation.Runtime" Version="1.*" />
```

### Web
```xml
<PackageReference Include="CsvHelper" Version="33.*" />
<PackageReference Include="DocumentFormat.OpenXml" Version="3.*" />
<PackageReference Include="MudBlazor" Version="7.*" />
<ProjectReference Include="..\EtikettenGenerator.ServiceDefaults\EtikettenGenerator.ServiceDefaults.csproj" />
```

### Tests
```xml
<PackageReference Include="xunit" Version="2.*" />
<PackageReference Include="xunit.runner.visualstudio" Version="2.*" />
<PackageReference Include="AwesomeAssertions" Version="*" />
<PackageReference Include="Microsoft.NET.Test.Sdk" Version="*" />
```

---

## C# Conventions

- **C# 13** features enabled
- `<Nullable>enable</Nullable>` in all projects
- `<ImplicitUsings>enable</ImplicitUsings>` in all projects
- `record` + `init` for immutable model types
- `sealed` on service classes and records where applicable
- No FluentAssertions — use **AwesomeAssertions** in all test projects

---

## MudBlazor Theme (CI Colors)

Configure in `App.razor` or `MainLayout.razor`:

```csharp
private readonly MudTheme _ciTheme = new()
{
    PaletteLight = new PaletteLight
    {
        Primary = "#000548",        // --ci-dark-blue
        Secondary = "#4a6fb8",      // --ci-blue
        Error = "#eb003c",          // --ci-red
        Success = "#2f8f68",        // --ci-green
        Warning = "#dee100",        // --ci-yellow
        GrayLight = "#c7ccd9",      // --ci-light-grey
        Surface = "#ffffff",        // --ci-white
        Background = "#ffffff",
    }
};
```

Wrap root layout with `<MudThemeProvider Theme="_ciTheme" />`.

### CI Color Reference

| Token              | Hex       | Usage               |
|--------------------|-----------|---------------------|
| `--ci-dark-blue`   | `#000548` | Primary / branding  |
| `--ci-red`         | `#eb003c` | Error / danger      |
| `--ci-blue`        | `#4a6fb8` | Secondary / links   |
| `--ci-green`       | `#2f8f68` | Success / confirm   |
| `--ci-yellow`      | `#dee100` | Warning / highlight |
| `--ci-light-grey`  | `#c7ccd9` | Backgrounds / lines |
| `--ci-white`       | `#ffffff` | Surface             |

---

## Architecture: Simple Layered (within Web project)

No external layers or separate assemblies beyond the Aspire structure.
Organize within `EtikettenGenerator.Web` as:

```
Models/       → data records (Member.cs)
Services/     → business logic (CsvImportService, WordExportService)
Components/   → reusable Blazor components
Pages/        → routable pages (Index.razor)
Templates/    → embedded Word template file
wwwroot/      → app.js (download helper)
```

Services are registered as **Scoped** (Blazor Server session-scoped).

---

## Key Domain Rules

- **CSV delimiter:** `;` (semicolon), auto-detect `,` as fallback
- **CSV encoding:** UTF-8 with BOM; fallback to Latin-1 (`ISO-8859-1`)
- **Missing columns:** throw `CsvImportException` with list of missing column names
- **Rettungsdienstfortbildung detection:** `OrdinalIgnoreCase` substring match on `"Rettungsdienstfortbildung"` in the `Ausbildungen` column
- **Word template placeholders:** `{{NACHNAME}}`, `{{VORNAME}}`, `{{MED_QUAL}}`, `{{DIENSTSTELLUNG}}`, `{{FAHRERLAUBNIS}}`, `{{RD_FORTBILDUNG}}`
- **Run-merge required** before placeholder replacement (placeholders may span multiple XML runs)
- **Overflow > 40 members:** append copies of the template table with page break
- **`HatRettungsdienstfortbildung: true`** → `"Ja"` / `false` → `""` (empty, not negative text)
- **No persistent storage:** all data lives in Scoped DI services only

---

## UI Behaviour Rules

- Upload: Blazor native `<InputFile>`, no JS interop for upload
- Loading spinner (`isLoading` state) during CSV parse and Word export
- Errors displayed in `MudAlert` (Severity.Error) above the table
- "Alle auswählen" checkbox: tri-state (all / none / mixed)
- **Full selection:** export immediately, no dialog
- **Partial selection:** open `PositionPickerDialog` (4×10 grid)
- **[Bestätigen]** in dialog only active when `selectedPositions.Count == selectedMembers.Count`
- Download via `IJSRuntime` → `window.downloadFile(filename, contentType, base64)`

---

## Observability

Add custom ActivitySource telemetry:

```csharp
// CsvImportService
using var activity = ActivitySource.StartActivity("csv.import");
activity?.SetTag("csv.rows", members.Count);

// WordExportService
using var activity = ActivitySource.StartActivity("word.export");
activity?.SetTag("export.members", members.Count);
activity?.SetTag("export.pages", pageCount);
```

---

## Testing Guidelines

- Framework: **xUnit**
- Assertions: **AwesomeAssertions** (never FluentAssertions)
- Unit test focus: `CsvImportService` and `WordExportService` logic
- Use `MemoryStream` for Word export tests (no file I/O)
- Test CSV edge cases: missing columns, empty file, BOM encoding, comma delimiter

---

## Getting Started

```bash
# Start the solution
dotnet run --project EtikettenGenerator.AppHost

# Or with Aspire CLI
aspire run

# Run tests
dotnet test
```

Aspire Dashboard opens automatically — shows logs, traces, metrics, and health status for the Web app.
