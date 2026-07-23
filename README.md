# Invoke-FileNameTransformation

`Invoke-FileNameTransformation` is an enterprise GUI-based PowerShell utility for advanced file renaming, structural transformation, and corporate naming compliance auditing. It helps users restructure filenames, enrich them with CSV/JSON/XML dictionary mappings, extract COM/EXIF file metadata, and preview all changes safely before copying or moving files.

---

## Key Features

- **7-Tab WPF Graphical Interface**:
  - **Tab 0: Naming Compliance Audit** (`Tab_Compliance`): Audit and automatically standardize filenames according to corporate conventions (`YYYYMMDD_XxxxxxxY_DocType_FreeText_v1.ext`), extract author and creation dates from COM file properties or EXIF metadata, and rename non-compliant files in place.
  - **Tab 1: Saved Profiles** (`Tab_Profile`): Save, load, duplicate, and manage reusable transformation profile configurations (JSON Schema V2).
  - **Tab 2: File Analysis** (`Tab_Analysis`): Group source files by token structure signature, adjust tokenizer regex patterns, and inspect sample filenames.
  - **Tab 3: Fields & Mappings** (`Tab_Fields`): Configure detected fields, resolve ambiguous types, attach per-field transformation chains (substring, date formatting, regex replacement, PowerShell scripts, sequential counters, casing, padding, math), and define dictionary lookups (CSV, TXT, JSON, XML).
  - **Tab 4: Destination Name Builder** (`Tab_DestName`): Interactively assemble target output filename templates using field tokens, virtual metadata fields (`File_Date`, `File_Author`, `Hash_MD5`, etc.), static text, and separators.
  - **Tab 5: Preview & Execution** (`Tab_Preview`): Generate a full transformation preview grid, detect collision errors or invalid path characters, export audit logs to CSV, and execute copy or move operations with real-time progress tracking.
  - **Tab 6: Custom Type Rules** (`Tab_CustomRules`): Define, edit, save, and live-test custom domain-specific regular expression recognition rules (e.g., department codes, invoice numbers, ticket IDs).
- **Interactive Tooltips**: Built-in, localized hover tooltips (`{t:ToolTip_...}`) across all UI buttons, inputs, drop-downs, and data grids in Polish, English, and German.
- **Typed Pattern Parsing**: Break incoming filenames into fields and automatically recognize text, integers, decimals, dates/times, GUIDs, versions, and configured custom types.
- **Composite Values**: Keep structured values such as `2026-01-16` or a punctuated GUID together even when their punctuation matches a filename separator.
- **Ambiguity Control**: Show competing interpretations and require an explicit field type before processing when automatic detection is not conclusive.
- **Data Mapping**: Connect external CSV, JSON, or XML dictionaries to replace extracted values with richer output data.
- **Metadata Integration**: Extract creation dates, author names, document titles, camera details, EXIF date taken, audio artists, and MD5/SHA256 file hashes.
- **Multi-language Interface**: Built-in support for Polish (PL), English (EN), and German (DE) with dynamic runtime switching.
- **Single-Executable Compilation**: Package the entire application—including core module, WPF XAML layout, localization dictionaries, and helper scripts—into a standalone EXE via `ps2exe.ps1`.

---

## System Requirements

- **Operating System**: Windows 10, Windows 11, or Windows Server 2016+
- **PowerShell**: Windows PowerShell 5.1 (or PowerShell 7.x with Windows Desktop SDK)
- **Dependencies**: Built-in WPF (`PresentationFramework`) and WinForms assemblies (`System.Windows.Forms`, `System.Drawing`).

---

## Getting Started

### Running the Development Script

Run the main script from PowerShell in the repository folder:

```powershell
.\Invoke-FileNameTransformation.ps1
```

For launch from Command Prompt, desktop shortcuts, or batch scripts:

```cmd
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\Invoke-FileNameTransformation.ps1"
```

> [!NOTE]
> The development launcher runs the GUI inside an isolated STA child process (`-IsolatedHost`) on every invocation. This permits closing the window, switching languages, and re-launching from the same PowerShell console without WPF thread locks or stale memory state.

### Using In-Command Help

Access full comment-based PowerShell help for all scripts and core functions:

```powershell
Get-Help .\Invoke-FileNameTransformation.ps1 -Full
Get-Help .\FileNameTransformation.Core.psm1 -Full
Get-Help .\ps2exe.ps1 -Full
Get-Help .\Setup-CustomTypeDemo.ps1 -Full
```

---

## Compiling to Standalone Executable

To compile the application into a single self-contained executable:

```powershell
# Build production executable (Invoke-FileNameTransformation.exe)
.\ps2exe.ps1 -Output "Invoke-FileNameTransformation.exe"

# Build debug executable with embedded console window
.\ps2exe.ps1 -Output "Invoke-FileNameTransformation_Debug.exe" -DebugBuild
```

The packager embeds:
- `FileNameTransformation.Core.psm1`
- `MainWindow.xaml`
- `locales/en.json`, `locales/pl.json`, `locales/de.json`
- All helper scripts in `src/*.ps1` (`Analysis.ps1`, `Compliance.ps1`, `Events.ps1`, `Mappings.ps1`, `Preview.ps1`, `Profiles.ps1`, `Transforms.ps1`, `Translations.ps1`, `UI.ps1`)

---

## Corporate Naming Compliance Audit

**Tab 0 (Compliance Audit)** implements automated auditing and fixing for enterprise naming conventions:

- **Target Standard**: `YYYYMMDD_XxxxxxxY_DocType_FreeText_v1.ext`
  - `YYYYMMDD`: 8-digit date string.
  - `XxxxxxxY`: 8-character PascalCase author code starting and ending with uppercase letters.
  - `DocType`: Document type classification keyword (e.g. `Raport`, `Umowa`, `Dokument`).
  - `FreeText`: Sanitized description text without invalid characters.
  - `v1`: Version tag (e.g. `v1`, `v2`, `v1.2`).
- **Metadata Fallbacks**: Automatically pulls creation dates and author names from Shell COM properties (`System.Author`), EXIF metadata, or file system timestamps.
- **In-Place Fixes**: Click **Fix Selected** (`Btn_FixSelected`) to safely rename non-compliant files in place directly within the source folder after review.

---

## Custom Data Type Rules

Custom named regex types can be managed via **Tab 6 (Custom Rules)** in the GUI or configured directly in `config.json`. For detailed documentation, schema details, and practical examples, see the [Custom Type Rules Guide](CUSTOM_TYPE_RULES.md).

```json
{
  "Version": 2,
  "Language": "EN",
  "CustomTypeRules": [
    {
      "Id": "DeptCode",
      "DisplayName": "Department Code",
      "Pattern": "^[A-Z]{3}-\\d{3}$",
      "Enabled": true,
      "AllowComposite": true
    }
  ]
}
```

To set up an interactive testing environment with sample files and pre-configured custom types:

```powershell
.\Setup-CustomTypeDemo.ps1
```

---

## Running Unit Tests

Execute the automated Pester test suite for tokenization, pattern matching, type inference, profile normalization, and compliance validation:

```powershell
Invoke-Pester .\tests\FileNameTransformation.Core.Tests.ps1
```

---

## License & Support

Developed for enterprise PowerShell file automation workflows. Designed by AI Pair Programming in collaboration with the Google DeepMind team.
