# Custom Type Rules Guide

This guide explains how to define, configure, and use **Custom Type Rules** in `Invoke-FileNameTransformation`.

Custom Type Rules extend the application's filename tokenizer and parser by recognizing domain-specific data types (such as department codes, project identifiers, ticket numbers, or custom serial numbers) alongside built-in types (`Integer`, `Decimal`, `DateTime`, `Guid`, `Version`).

---

## Table of Contents

- [Overview](#overview)
- [Rule Structure & Schema](#rule-structure--schema)
- [Configuring Custom Type Rules](#configuring-custom-type-rules)
  - [Method 1: Using the GUI (Tab 6: Custom Type Rules)](#method-1-using-the-gui-tab-6-custom-type-rules)
  - [Method 2: Manual JSON Configuration (`config.json`)](#method-2-manual-json-configuration-configjson)
- [How Custom Rules Work](#how-custom-rules-work)
  - [Regex Anchoring & Full-Value Matching](#regex-anchoring--full-value-matching)
  - [Composite Token Matching (`AllowComposite`)](#composite-token-matching-allowcomposite)
- [Practical Examples](#practical-examples)
  - [Example 1: Department Code](#example-1-department-code)
  - [Example 2: Project Identifier](#example-2-project-identifier)
  - [Example 3: Invoice Number](#example-3-invoice-number)
  - [Example 4: Service Desk Ticket ID](#example-4-service-desk-ticket-id)
  - [Example 5: Hardware Serial Number](#example-5-hardware-serial-number)
  - [Example 6: SKU Code with Version](#example-6-sku-code-with-version)
- [Testing Custom Types (Interactive Demo Script)](#testing-custom-types-interactive-demo-script)
- [Troubleshooting & Common Errors](#troubleshooting--common-errors)

---

## Overview

When analyzing filenames, `Invoke-FileNameTransformation` breaks filename strings into discrete segments (tokens) separated by delimiters such as underscores (`_`), hyphens (`-`), or dots (`.`).

By default, the engine recognizes five standard built-in types:
- **Integer**: Whole numbers (e.g., `123`, `-45`)
- **Decimal**: Floating-point numbers (e.g., `123.45`)
- **DateTime**: Standard date and time formats (e.g., `2026-01-16`, `010225`)
- **Guid**: 128-bit Globally Unique Identifiers (e.g., `6F9619FF-8B86-D011-B42D-00C04FC964FF`)
- **Version**: Version numbers with 2 to 4 components (e.g., `1.2.3`, `2.0.0.1`)

**Custom Type Rules** allow you to define custom named regular expression (regex) patterns. When a filename segment matches your custom rule, the application classifies it under your custom type name (e.g., `Custom:DeptCode`).

---

## Rule Structure & Schema

A Custom Type Rule object consists of five fields:

| Field | Data Type | Required? | Default Value | Description |
| :--- | :--- | :--- | :--- | :--- |
| `Id` | String | **Yes** | — | Unique identifier for the rule (e.g., `DeptCode`). Must start with a letter and contain only letters, numbers, dots, hyphens, or underscores (`^[A-Za-z][A-Za-z0-9_.-]*$`). Case-insensitively unique. |
| `DisplayName` | String | No | Defaults to `Id` | User-friendly label displayed in the UI and type dropdowns (e.g., `"Department Code"`). |
| `Pattern` | String | **Yes** | — | A valid .NET Regular Expression pattern used to evaluate values (e.g., `^[A-Z]{3}-\d{3}$`). |
| `Enabled` | Boolean | No | `true` | When `true`, the rule is active and used during token analysis. When `false`, the rule is preserved in configuration but ignored during analysis. |
| `AllowComposite` | Boolean | No | `false` | When `true`, allows the custom rule to match multi-segment composite values spanning across tokenizer separators. |

---

## Configuring Custom Type Rules

You can manage Custom Type Rules using either the graphical interface or by editing `config.json`.

### Method 1: Using the GUI (Tab 6: Custom Type Rules)

1. **Launch the Application**: Run `Invoke-FileNameTransformation.ps1` or `Invoke-FileNameTransformation.exe`.
2. **Open the Custom Rules Tab**: Click on **Tab 6: Custom Rules Management & Live Tester** (or *"Custom Type Rules"* / *"Zasady typów własnych"* depending on language).
3. **Add or Edit a Rule**:
   - Click **New Rule** to populate the fields with a template.
   - Enter the **Rule ID** (e.g., `DeptCode`).
   - Enter a human-readable **Display Name** (e.g., `Department Code`).
   - Enter your **Regex Pattern** (e.g., `^[A-Z]{3}-\d{3}$`).
   - Toggle **Enabled** and **Allow Composite** checkboxes as needed.
   - Click **Apply / Update Rule** to save the rule to the active session.
4. **Test the Rule Live**:
   - In the **Live Custom Rule Tester** box, enter a test string (e.g., `FIN-101`).
   - Click **Test Value**. The test results will list all candidate types matched by the engine.
5. **Persist Settings**:
   - Click **Save Rules to Config** to save all configured rules permanently to `config.json`.

---

### Method 2: Manual JSON Configuration (`config.json`)

You can edit `config.json` directly in the application directory.

> [!IMPORTANT]
> When writing Regex in JSON files, you **must escape backslashes** (e.g., write `\\d` instead of `\d`, `\\w` instead of `\w`).

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
      "AllowComposite": false
    },
    {
      "Id": "InvoiceNo",
      "DisplayName": "Invoice Number",
      "Pattern": "^INV-\\d{6}-\\d{5}$",
      "Enabled": true,
      "AllowComposite": true
    }
  ]
}
```

---

## How Custom Rules Work

### Regex Anchoring & Full-Value Matching

The engine requires the regex pattern to match the **entire candidate value**.
For instance, the pattern `[A-Z]{3}` will match `ABC`, but if the token value is `ABC-123`, matching only `ABC` is insufficient unless anchored properly or matched as a single token.

**Best Practice**: Always anchor your regular expression patterns using `^` (start of string) and `$` (end of string):
- **Correct**: `^[A-Z]{3}-\d{3}$`
- **Avoid**: `[A-Z]{3}-\d{3}` (unanchored patterns may cause unexpected partial matches during token evaluation)

### Composite Token Matching (`AllowComposite`)

By default, the tokenizer splits filenames at standard delimiters (like `-` or `_`).

For example, given the filename segment `FIN-101`:
- If `AllowComposite` is `false`, the engine might tokenize `FIN` and `101` as two separate fields.
- If `AllowComposite` is `true`, the engine evaluates `FIN-101` as a single composite token against custom rules that support composite matching.

---

## Practical Examples

Here are 6 practical examples of Custom Type Rules for common file naming conventions:

### Example 1: Department Code
Matches standard 3-letter department codes followed by a hyphen and 3 digits (e.g., `FIN-101`, `HRD-202`, `ITS-305`).

- **Id**: `DeptCode`
- **DisplayName**: `Department Code`
- **Pattern**: `^[A-Z]{3}-\d{3}$`
- **AllowComposite**: `true`
- **Example Filename**: `Report_FIN-101_2026-01-16.pdf`

```json
{
  "Id": "DeptCode",
  "DisplayName": "Department Code",
  "Pattern": "^[A-Z]{3}-\\d{3}$",
  "Enabled": true,
  "AllowComposite": true
}
```

---

### Example 2: Project Identifier
Matches project codes consisting of `PRJ` or `PROJECT`, a 4-digit year, and a 3-digit sequence (e.g., `PRJ-2026-042`, `PROJECT-2025-001`).

- **Id**: `ProjectID`
- **DisplayName**: `Project Identifier`
- **Pattern**: `^PRJ(ECT)?-\d{4}-\d{3}$`
- **AllowComposite**: `true`
- **Example Filename**: `Doc_PRJ-2026-042_Draft.docx`

```json
{
  "Id": "ProjectID",
  "DisplayName": "Project Identifier",
  "Pattern": "^PRJ(ECT)?-\\d{4}-\\d{3}$",
  "Enabled": true,
  "AllowComposite": true
}
```

---

### Example 3: Invoice Number
Matches corporate invoice numbers starting with `INV-`, a 6-digit YYYYMM date block, a hyphen, and a 5-digit invoice ID (e.g., `INV-202607-00123`).

- **Id**: `InvoiceNo`
- **DisplayName**: `Invoice Number`
- **Pattern**: `^INV-\d{6}-\d{5}$`
- **AllowComposite**: `true`
- **Example Filename**: `INV-202607-00123_AcmeCorp.pdf`

```json
{
  "Id": "InvoiceNo",
  "DisplayName": "Invoice Number",
  "Pattern": "^INV-\\d{6}-\\d{5}$",
  "Enabled": true,
  "AllowComposite": true
}
```

---

### Example 4: Service Desk Ticket ID
Matches IT Service Management prefixes (`INC` for Incident, `TASK` for Task, `REQ` for Request) followed by 6 digits (e.g., `INC123456`, `TASK998877`, `REQ001122`).

- **Id**: `TicketID`
- **DisplayName**: `Service Desk Ticket`
- **Pattern**: `^(INC|TASK|REQ)\d{6}$`
- **AllowComposite**: `false`
- **Example Filename**: `INC123456_LogFile_2026.txt`

```json
{
  "Id": "TicketID",
  "DisplayName": "Service Desk Ticket",
  "Pattern": "^(INC|TASK|REQ)\\d{6}$",
  "Enabled": true,
  "AllowComposite": false
}
```

---

### Example 5: Hardware Serial Number
Matches serial numbers beginning with `SN-`, followed by 4 hexadecimal characters and 4 digits (e.g., `SN-A8F9-2026`).

- **Id**: `SerialNumber`
- **DisplayName**: `Serial Number`
- **Pattern**: `^SN-[A-F0-9]{4}-\d{4}$`
- **AllowComposite**: `true`
- **Example Filename**: `Diagnostics_SN-A8F9-2026.log`

```json
{
  "Id": "SerialNumber",
  "DisplayName": "Serial Number",
  "Pattern": "^SN-[A-F0-9]{4}-\\d{4}$",
  "Enabled": true,
  "AllowComposite": true
}
```

---

### Example 6: SKU Code with Version
Matches product stock keeping units with revision numbers (e.g., `SKU-AB12-v1`, `SKU-XY99-V3`).

- **Id**: `SKUCode`
- **DisplayName**: `Product SKU`
- **Pattern**: `^SKU-[A-Z]{2}\d{2}-[vV]\d+$`
- **AllowComposite**: `true`
- **Example Filename**: `Spec_SKU-AB12-v1_Approved.pdf`

```json
{
  "Id": "SKUCode",
  "DisplayName": "Product SKU",
  "Pattern": "^SKU-[A-Z]{2}\\d{2}-[vV]\\d+$",
  "Enabled": true,
  "AllowComposite": true
}
```

---

## Testing Custom Types (Interactive Demo Script)

To quickly test Custom Type Rules without creating test files manually, you can run the included setup script [`Setup-CustomTypeDemo.ps1`](file:///d:/Skrypty/Invoke-FileNameTransformation/Setup-CustomTypeDemo.ps1).

### Quick Start Command

Run the setup script in PowerShell:

```powershell
powershell -ExecutionPolicy Bypass -File .\Setup-CustomTypeDemo.ps1
```

### What the Setup Script Does

1. **Creates Demo Directory & Test Files**: Generates a `Demo_CustomTypes` folder populated with sample files featuring realistic custom type naming formats (Department Codes, Invoices, Service Desk Tickets, Project IDs, and Hardware Serial Numbers).
2. **Pre-configures `config.json`**: Automatically adds sample Custom Type Rules to `config.json` (and backs up any existing `config.json` to `config.json.bak`).
3. **Generates Quick-Start Guide**: Creates a localized `README_TESTING.md` inside `Demo_CustomTypes` for easy reference during testing.

### End-to-End Testing Walkthrough

Once the script completes, follow these steps to test custom types in the tool:

1. **Launch the Application**:
   Run `Invoke-FileNameTransformation.ps1` in PowerShell or double-click `Invoke-FileNameTransformation.exe`.

2. **Verify Loaded Rules (Tab 6)**:
   - Click on **Tab 6: Custom Rules Management & Live Tester**.
   - Review the newly configured custom type rules (`DeptCode`, `InvoiceNo`, `TicketID`, `ProjectID`, `HardwareSN`).
   - Use the **Live Custom Rule Tester** to test sample inputs like `FIN-101` or `INV-202607-00123`.

3. **Select the Demo Folder & Run Analysis (Tab 1)**:
   - Go to **Tab 1: File Selection & Metadata**.
   - Click **Browse** and select the generated `Demo_CustomTypes` folder.
   - Click **Run Token Analysis**.
   - The application will break down each filename and tag recognized tokens with types like `Custom:DeptCode`, `Custom:InvoiceNo`, `Custom:TicketID`, `DateTime`, etc.

4. **Define Target Naming Convention (Tab 2)**:
   - Go to **Tab 2: Transformation Rules**.
   - Build a target naming pattern placing custom tokens in desired positions (e.g., rearrange `Custom:DeptCode` to field position 1).

5. **Preview & Apply Transformation (Tab 3)**:
   - Go to **Tab 3: Preview & Execute**.
   - Click **Generate Preview** to review the proposed filename transformations.
   - Click **Apply Transformation** to perform the file renaming.

---

## Troubleshooting & Common Errors

At application startup, `Invoke-FileNameTransformation` validates all custom type rules. If a rule fails validation, a warning is displayed and the broken rule is skipped without crashing the application.

### Common Validation Issues

| Error Message Pattern | Cause | Solution |
| :--- | :--- | :--- |
| `Rule X has an invalid ID '123bad'` | Rule `Id` starts with a number or contains invalid characters. | Change `Id` so it starts with an ASCII letter (`A-Z`, `a-z`) and contains only letters, numbers, dots, hyphens, or underscores. |
| `Rule X duplicates ID 'deptcode'` | Multiple rules share the same `Id` (case-insensitive). | Ensure every rule has a unique `Id`. |
| `Rule 'X' has an empty pattern.` | `Pattern` field is missing or empty. | Provide a non-empty regex pattern string. |
| `Rule 'X' has an invalid regex: ...` | Syntax error in the regular expression. | Verify your regex syntax (e.g. check for unclosed brackets `[` or invalid escape sequences). If editing `config.json`, make sure backslashes are doubled (`\\d`). |

### JSON Backslash Escaping Tip

If your rule works in the GUI Live Tester but fails when loaded from `config.json`, check your backslashes:

- **GUI Input Field**: Enter standard single backslashes (e.g., `^[A-Z]{3}-\d{3}$`).
- **`config.json` File**: Enter doubled backslashes (e.g., `"Pattern": "^[A-Z]{3}-\\d{3}$"`).
