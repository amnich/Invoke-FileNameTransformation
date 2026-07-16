# Invoke-FileNameTransformation

Invoke-FileNameTransformation is a GUI-based PowerShell utility for advanced file renaming and transformation. It helps users restructure filenames, enrich them with CSV-based mappings, and preview all changes safely before copying or moving files.

## Features

- **Typed Pattern Parsing**: Break incoming filenames into fields and automatically recognize text, integers, decimals, dates/times, GUIDs, versions, and configured custom types.
- **Composite Values**: Keep structured values such as `2026-01-16` or a punctuated GUID together even when their punctuation is also a filename separator.
- **Ambiguity Control**: Show competing interpretations and require an explicit field type before processing when automatic detection is not conclusive.
- **Data Mapping**: Connect external CSV dictionaries to replace extracted values with richer output data.
- **Text Transformations**: Apply substring extraction, padding, casing changes, replacements, and date formatting to any part of the filename.
- **Output Templates**: Build target filenames from static text, extracted fields, and mapped values.
- **Live Preview & Validation**: Preview the resulting file names in a grid and detect collisions, missing values, or invalid characters before execution.
- **Saveable Profiles**: Export and reload complex configurations as JSON profiles.
- **Multi-language Interface**: Built-in support for Polish, English, and German.
- **Copy or Move Execution**: Run the transformation in copy or move mode and keep an audit log.
- **Folder Layout Control**: Scan source subfolders independently from preserving their layout in the destination folder.

## Requirements

- Windows PowerShell 5.1
- Windows OS with WPF and WinForms support
- `FileNameTransformation.Core.psm1` in the same folder as the script or packaged executable

## Usage

Run the script from PowerShell in the repository folder:

```powershell
.\Invoke-FileNameTransformation.ps1
```

For a command prompt, shortcut, or automation scenario, start it explicitly in STA mode:

```cmd
powershell.exe -NoProfile -STA -File ".\Invoke-FileNameTransformation.ps1"
```

## Typical Workflow

1. **Select folders**: Choose the source folder and output destination. Enable **Scan subfolders** to include nested files. Use **Preserve folder structure in destination folder** to recreate the source hierarchy at the destination, or clear it to copy or move all results directly into the selected destination folder.
2. **Analyze names**: Review the detected filename structures and choose a pattern.
3. **Define fields and mappings**: Review detected data types, resolve fields marked as ambiguous, apply transforms, and load CSV lookup files.
4. **Build the output name**: Assemble the target filename from fields, text, and separators.
5. **Preview and execute**: Review the results and run the copy or move operation.

## Tokenizer Regex Pattern

The **Regex Pattern** determines how each filename is split into values and separators during analysis. The expression must match every segment of the filename. Use the named capture group `(?<sep>...)` for literal separators; every other match is treated as a value that can become a field.

The default pattern splits on underscores, hyphens, and whitespace:

```regex
(?<value>[^_\-\s]+)|(?<sep>[_\-\s]+)
```

The expression is validated before analysis. It must define `sep`, cannot produce zero-length matches, and must cover the complete filename. For example, `Report_20260116-Final` initially becomes the values `Report`, `20260116`, and `Final`, with `_` and `-` preserved as separators. To use different separators, include them in the `sep` group, such as `(?<value>[^.]+)|(?<sep>[.])` for dot-separated names.

After lexical tokenization, semantic recognition evaluates individual values and adjacent spans. This allows `Report_2025-01-16_Final` to contain one date field even though hyphens are normal separators. Recognition uses invariant numeric parsing and calendar-valid exact date formats instead of only matching the value's shape.

## Data Types and Ambiguity

Built-in recognition supports:

- Text
- Integer
- Decimal using the invariant `.` decimal separator
- Date and time in supported exact formats
- GUID
- Version with two to four numeric components

When a value has several valid meanings, the field is marked **Ambiguous**. For example, `010225` can be an integer and can also match multiple compact date formats. Select the field and choose its intended data type and date format before building the full preview. The selection is stored in version 2 profiles.

Pattern enforcement is intentionally strict. Other files must have the same number and order of fields and the same literal separators. Preview errors identify the failing token, expected structure or type, and actual value; the application does not guess missing fields.

## Custom Data Types

Custom named regex types are configured in `config.json`. Patterns must match the complete candidate value. Set `AllowComposite` to allow a rule to join values across tokenizer separators.

```json
{
  "Version": 2,
  "Language": "EN",
  "CustomTypeRules": [
    {
      "Id": "DepartmentCode",
      "Pattern": "^[A-Z]{3}-\\d{3}$",
      "Enabled": true,
      "AllowComposite": true
    }
  ]
}
```

Changing the UI language preserves custom rules. Invalid custom patterns are reported during analysis.

## Tests

Run the parser and matcher tests with Pester:

```powershell
Invoke-Pester .\tests\FileNameTransformation.Core.Tests.ps1
```

## Notes

- Profiles, logs, and the language configuration are stored under the user's AppData folder.
- Existing profiles without a schema version are migrated in memory when loaded. They are not overwritten automatically.
- If AppData is unavailable, the script falls back to the script directory or a temporary folder.
- Preserving the folder structure is enabled by default. When it is disabled, files from different subfolders can produce duplicate destination names; the preview flags those collisions before execution.
- Keep `FileNameTransformation.Core.psm1` beside `Invoke-FileNameTransformation.exe` when distributing the packaged application.
