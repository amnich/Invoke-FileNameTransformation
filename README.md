# FileNameTransformer

FileNameTransformer is a powerful GUI-based PowerShell utility for advanced file renaming and transformation. It enables users to perform structural filename modifications, utilize mappings from CSV files, and preview all changes safely before copying or renaming files.

## Features

- **Pattern Parsing**: Automatically break down incoming filenames into structured parts (e.g. `[Prefix]-[ID]-[Name]`).
- **Data Mapping**: Connect external CSV dictionaries to map extracted values to new outputs (e.g. converting a country code into a full country name).
- **Text Transformations**: Apply regex replacements, substring extractions, padding, casing modifications, and date formatting directly to any part of the filename.
- **Advanced Output Templates**: Construct your target filenames flexibly by mixing static text, extracted properties, and mapped data.
- **Live Preview & Validation**: Safely preview your renaming operations in a data grid. The application verifies constraints, checks for potential file collisions, and flags missing data elements prior to execution.
- **Saveable Profiles**: Export and load your complex mapping configurations as JSON profiles for seamless reuse.
- **Multi-language Interface**: Built-in support for Polish, English, and German interfaces.

## Requirements

- Windows PowerShell 5.1
- Windows OS (Requires WPF and WinForms capability)

## Usage

You can launch the tool simply by running the script in PowerShell. It requires STA (Single-Threaded Apartment) mode, which is standard for WPF-based scripts.

```powershell
.\FileNameTransformer.GUI.ps1
```

If you are invoking it from a command prompt or shortcut, you should force STA mode:

```cmd
powershell.exe -STA -File ".\FileNameTransformer.GUI.ps1"
```

## How It Works

1. **Source Selection**: Pick the folder containing the files you want to process, and select an output destination.
2. **Define Patterns**: Specify how the source filenames are currently structured using variables.
3. **Map and Transform**: Define virtual fields that transform data, or use mapping tables from external CSVs.
4. **Construct Template**: Build your output filename template using the defined fields.
5. **Execute**: Review the live preview grid and execute the renaming operation. The application copies the files with their new names.
