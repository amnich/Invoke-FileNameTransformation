# Custom Type Rules - Testing Guide

Welcome! This folder contains test files pre-configured to work with **Invoke-FileNameTransformation**.

## Configured Custom Types

The following custom data types have been defined in config.json:

| Custom Type ID | Display Name | Pattern / Format Example | Example Filename |
| :--- | :--- | :--- | :--- |
| DeptCode | Department Code | 3 letters, hyphen, 3 numbers (FIN-101, HRD-202) | Report_FIN-101_2026-01-16.pdf |
| InvoiceNo | Invoice Number | INV- + 6 digits + - + 5 digits (INV-202607-00123) | INV-202607-00123_AcmeCorp.pdf |
| TicketID | Service Ticket ID | TICKET- + 5 digits (TICKET-84920) | TICKET-84920_Urgent_Fix.docx |
| ProjectID | Project Identifier | PRJ-YYYY-XXX or PROJECT-YYYY-XXX (PRJ-2026-042) | Doc_PRJ-2026-042_Draft.docx |
| HardwareSN | Serial Number | SN-XXXX-XXXX (SN-A1B2-C3D4) | Audit_SN-A1B2-C3D4_Server01.log |

---

## How to Test Custom Types (Step-by-Step)

### 1. Launch the Application
Run Invoke-FileNameTransformation.ps1 in PowerShell, or execute Invoke-FileNameTransformation.exe.

### 2. View/Manage Custom Types (Optional)
- Switch to **Tab 6: Custom Rules Management & Live Tester**.
- You will see the configured rules (DeptCode, InvoiceNo, TicketID, etc.).
- Try typing FIN-101 in the **Live Custom Rule Tester** box and click **Test Value** to verify matching.

### 3. Select Demo Folder & Analyze Files
- Switch to **Tab 1: File Selection & Metadata**.
- Click **Browse** and select this $TargetFolderName folder.
- Click **Run Token Analysis**.
- Notice how the application automatically identifies filename tokens and tags them with types like Custom:DeptCode, Custom:InvoiceNo, DateTime, Integer, etc.

### 4. Build a Transformation Rule
- Go to **Tab 2: Transformation Rules**.
- Choose a target naming convention (e.g. rearrange fields, add prefixes, or change casing).
- Map custom type tokens (e.g., place Custom:DeptCode at position 1, DateTime at position 2).

### 5. Preview & Apply
- Click **Generate Preview** to review the proposed new filenames.
- Verify everything looks accurate and click **Apply Transformation** to rename the files.

---

## Defining Your Own Custom Types

You can add new custom types anytime:
1. **Via GUI**: Tab 6 -> Click **New Rule** -> Enter ID, Display Name, Pattern (anchored regex like ^[A-Z]{2}-\d{4}$), check **Allow Composite** if needed -> Click **Apply / Update Rule** -> Click **Save Rules to Config**.
2. **Via config.json**: Open config.json in a text editor and add a new rule object under "CustomTypeRules". Remember to escape backslashes (e.g., \\d).
