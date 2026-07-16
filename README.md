# Invoke-FileNameTransformation

Invoke-FileNameTransformation is a GUI-based PowerShell utility for advanced file renaming and transformation. It helps users restructure filenames, enrich them with CSV-based mappings, and preview all changes safely before copying or moving files.

## Features

- **Pattern Parsing**: Automatically break down incoming filenames into structured parts (for example, `[Prefix]-[ID]-[Name]`).
- **Data Mapping**: Connect external CSV dictionaries to replace extracted values with richer output data.
- **Text Transformations**: Apply substring extraction, padding, casing changes, replacements, and date formatting to any part of the filename.
- **Output Templates**: Build target filenames from static text, extracted fields, and mapped values.
- **Live Preview & Validation**: Preview the resulting file names in a grid and detect collisions, missing values, or invalid characters before execution.
- **Saveable Profiles**: Export and reload complex configurations as JSON profiles.
- **Multi-language Interface**: Built-in support for Polish, English, and German.
- **Copy or Move Execution**: Run the transformation in copy or move mode and keep an audit log.

## Requirements

- Windows PowerShell 5.1
- Windows OS with WPF and WinForms support

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

1. **Select folders**: Choose the source folder and output destination.
2. **Analyze names**: Review the detected filename structures and choose a pattern.
3. **Define fields and mappings**: Create fields, apply transforms, and load CSV lookup files.
4. **Build the output name**: Assemble the target filename from fields, text, and separators.
5. **Preview and execute**: Review the results and run the copy or move operation.

## Notes

- Profiles, logs, and the language configuration are stored under the user's AppData folder.
- If AppData is unavailable, the script falls back to the script directory or a temporary folder.
