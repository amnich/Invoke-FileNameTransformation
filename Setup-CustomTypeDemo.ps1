<#
.SYNOPSIS
    Creates a demo testing environment for Custom Type Rules in Invoke-FileNameTransformation.

.DESCRIPTION
    This script:
    1. Creates a sample folder 'Demo_CustomTypes' populated with realistic test files.
    2. Updates 'config.json' with demo Custom Type Rules (DeptCode, InvoiceNo, TicketID, ProjectID, HardwareSN).
    3. Generates a quick-start testing guide (README_TESTING.md) inside the demo folder.
    4. Displays step-by-step instructions for testing custom types in the tool.
#>

[CmdletBinding()]
param(
    [string]$TargetFolderName = "Demo_CustomTypes",
    [switch]$SkipConfigUpdate
)

$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$DemoFolderPath = Join-Path $ScriptDir $TargetFolderName
$ConfigPath = Join-Path $ScriptDir "config.json"

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  Invoke-FileNameTransformation - Custom Type Demo Setup   " -ForegroundColor White
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

# 1. Create Demo Directory & Files
if (-not (Test-Path -Path $DemoFolderPath)) {
    New-Item -Path $DemoFolderPath -ItemType Directory | Out-Null
    Write-Host "[+] Created demo directory: $DemoFolderPath" -ForegroundColor Green
}
else {
    Write-Host "[!] Demo directory already exists: $DemoFolderPath" -ForegroundColor Yellow
}

$sampleFiles = @(
    @{ Name = "Report_FIN-101_2026-01-16.pdf"; Content = "Financial report for Dept FIN-101" },
    @{ Name = "Report_HRD-202_2026-02-01.docx"; Content = "HR report for Dept HRD-202" },
    @{ Name = "Report_MKT-303_2026-03-15.pptx"; Content = "Marketing report for Dept MKT-303" },
    @{ Name = "Invoice_HRD-202_2026-02-01.xlsx"; Content = "HR Department Invoice" },
    @{ Name = "INV-202607-00123_AcmeCorp.pdf"; Content = "Invoice for Acme Corp" },
    @{ Name = "INV-999999-54321_CosmoCorp.pdf"; Content = "Invoice for Cosmo Corp" },
    @{ Name = "INV-202607-00124_GlobalTech.pdf"; Content = "Invoice for Global Tech" },
    @{ Name = "TICKET-84920_Urgent_Fix.docx"; Content = "Service desk ticket documentation" },
    @{ Name = "TICKET-10492_Database_Migration.log"; Content = "Migration log for ticket 10492" },
    @{ Name = "TICKET-00001_Initial_Setup.txt"; Content = "Initial setup instructions for ticket 00001" },
    @{ Name = "Doc_PRJ-2026-042_Draft.docx"; Content = "Project 2026-042 draft specs" },
    @{ Name = "Doc_PROJECT-2025-001_Final.pdf"; Content = "Project 2025-001 final document" },
    @{ Name = "Doc_PRJ-2026-042_ReviewNotes.txt"; Content = "Review notes for project 2026-042" },
    @{ Name = "Audit_SN-A1B2-C3D4_Server01.log"; Content = "Hardware audit log for serial SN-A1B2-C3D4" },
    @{ Name = "Audit_SN-Z9Y8-X7W6_Server02.log"; Content = "Hardware audit log for serial SN-Z9Y8-X7W6" },
    @{ Name = "Audit_SN-1234-5678_Server03.log"; Content = "Hardware audit log for serial SN-1234-5678" }
)

Write-Host "[+] Populating demo files..." -ForegroundColor Green
foreach ($file in $sampleFiles) {
    $filePath = Join-Path $DemoFolderPath $file.Name
    if (-not (Test-Path -Path $filePath)) {
        Set-Content -Path $filePath -Value $file.Content -Encoding UTF8
        Write-Host "    - Created: $($file.Name)" -ForegroundColor Gray
    }
    else {
        Write-Host "    - Skipped (already exists): $($file.Name)" -ForegroundColor DarkGray
    }
}

# 2. Define Custom Type Rules
$demoRules = @(
    [pscustomobject]@{
        Id             = "DeptCode"
        DisplayName    = "Department Code"
        Pattern        = '^[A-Z]{3}-\d{3}$'
        Enabled        = $true
        AllowComposite = $true
    },
    [pscustomobject]@{
        Id             = "InvoiceNo"
        DisplayName    = "Invoice Number"
        Pattern        = '^INV-\d{6}-\d{5}$'
        Enabled        = $true
        AllowComposite = $true
    },
    [pscustomobject]@{
        Id             = "TicketID"
        DisplayName    = "Service Desk Ticket ID"
        Pattern        = '^TICKET-\d{5}$'
        Enabled        = $true
        AllowComposite = $true
    },
    [pscustomobject]@{
        Id             = "ProjectID"
        DisplayName    = "Project Identifier"
        Pattern        = '^PRJ(ECT)?-\d{4}-\d{3}$'
        Enabled        = $true
        AllowComposite = $true
    },
    [pscustomobject]@{
        Id             = "HardwareSN"
        DisplayName    = "Hardware Serial Number"
        Pattern        = '^SN-[A-Z0-9]{4}-[A-Z0-9]{4}$'
        Enabled        = $true
        AllowComposite = $true
    }
)

# 3. Update config.json if requested
if (-not $SkipConfigUpdate) {
    if (Test-Path -Path $ConfigPath) {
        try {
            $rawJson = Get-Content -Path $ConfigPath -Raw -ErrorAction Stop
            $configObj = $rawJson | ConvertFrom-Json
        }
        catch {
            $configObj = [pscustomobject]@{ Version = 2; Language = "EN"; CustomTypeRules = @(); Theme = "Dark" }
        }

        # Backup existing config if no backup exists yet
        $backupPath = "$ConfigPath.bak"
        if (-not (Test-Path -Path $backupPath)) {
            Copy-Item -Path $ConfigPath -Destination $backupPath -Force
            Write-Host "[+] Backed up existing config to: $backupPath" -ForegroundColor Green
        }

        # Merge rules (avoid duplicate IDs)
        $existingRules = if ($configObj.CustomTypeRules) { [array]$configObj.CustomTypeRules } else { @() }
        $existingIds = $existingRules | ForEach-Object { $_.Id.ToString().ToUpperInvariant() }

        $rulesToAdd = @()
        foreach ($r in $demoRules) {
            if (-not ($existingIds -contains $r.Id.ToUpperInvariant())) {
                $rulesToAdd += $r
            }
        }

        $configObj.CustomTypeRules = $existingRules + $rulesToAdd
        $updatedJson = $configObj | ConvertTo-Json -Depth 10
        Set-Content -Path $ConfigPath -Value $updatedJson -Encoding UTF8
        Write-Host "[+] Updated '$ConfigPath' with demo custom type rules." -ForegroundColor Green
    }
}

# 4. Create README_TESTING.md inside the demo folder
$readmePath = Join-Path $DemoFolderPath "README_TESTING.md"
$readmeContent = @"
# Custom Type Rules - Testing Guide

Welcome! This folder contains test files pre-configured to work with **Invoke-FileNameTransformation**.

## Configured Custom Types

The following custom data types have been defined in `config.json`:

| Custom Type ID | Display Name | Pattern / Format Example | Example Filename |
| :--- | :--- | :--- | :--- |
| `DeptCode` | Department Code | 3 letters, hyphen, 3 numbers (`FIN-101`, `HRD-202`) | `Report_FIN-101_2026-01-16.pdf` |
| `InvoiceNo` | Invoice Number | `INV-` + 6 digits + `-` + 5 digits (`INV-202607-00123`) | `INV-202607-00123_AcmeCorp.pdf` |
| `TicketID` | Service Ticket ID | `TICKET-` + 5 digits (`TICKET-84920`) | `TICKET-84920_Urgent_Fix.docx` |
| `ProjectID` | Project Identifier | `PRJ-YYYY-XXX` or `PROJECT-YYYY-XXX` (`PRJ-2026-042`) | `Doc_PRJ-2026-042_Draft.docx` |
| `HardwareSN` | Serial Number | `SN-XXXX-XXXX` (`SN-A1B2-C3D4`) | `Audit_SN-A1B2-C3D4_Server01.log` |

---

## How to Test Custom Types (Step-by-Step)

### 1. Launch the Application
Run `Invoke-FileNameTransformation.ps1` in PowerShell, or execute `Invoke-FileNameTransformation.exe`.

### 2. View/Manage Custom Types (Optional)
- Switch to **Tab 6: Custom Rules Management & Live Tester**.
- You will see the configured rules (`DeptCode`, `InvoiceNo`, `TicketID`, etc.).
- Try typing `FIN-101` in the **Live Custom Rule Tester** box and click **Test Value** to verify matching.

### 3. Select Demo Folder & Analyze Files
- Switch to **Tab 1: File Selection & Metadata**.
- Click **Browse** and select this `$TargetFolderName` folder.
- Click **Run Token Analysis**.
- Notice how the application automatically identifies filename tokens and tags them with types like `Custom:DeptCode`, `Custom:InvoiceNo`, `DateTime`, `Integer`, etc.

### 4. Build a Transformation Rule
- Go to **Tab 2: Transformation Rules**.
- Choose a target naming convention (e.g. rearrange fields, add prefixes, or change casing).
- Map custom type tokens (e.g., place `Custom:DeptCode` at position 1, `DateTime` at position 2).

### 5. Preview & Apply
- Click **Generate Preview** to review the proposed new filenames.
- Verify everything looks accurate and click **Apply Transformation** to rename the files.

---

## Defining Your Own Custom Types

You can add new custom types anytime:
1. **Via GUI**: Tab 6 -> Click **New Rule** -> Enter ID, Display Name, Pattern (anchored regex like `^[A-Z]{2}-\d{4}$`), check **Allow Composite** if needed -> Click **Apply / Update Rule** -> Click **Save Rules to Config**.
2. **Via `config.json`**: Open `config.json` in a text editor and add a new rule object under `"CustomTypeRules"`. Remember to escape backslashes (e.g., `\\d`).
"@

Set-Content -Path $readmePath -Value $readmeContent -Encoding UTF8
Write-Host "[+] Created quick-start instructions: $readmePath" -ForegroundColor Green

# 5. Display Console Instructions
Write-Host ""
Write-Host "============================================================" -ForegroundColor Yellow
Write-Host "                  HOW TO TEST CUSTOM TYPES                  " -ForegroundColor Yellow
Write-Host "============================================================" -ForegroundColor Yellow
Write-Host "1. Launch the GUI tool:" -ForegroundColor White
Write-Host "   powershell -ExecutionPolicy Bypass -File .\Invoke-FileNameTransformation.ps1" -ForegroundColor Cyan
Write-Host "   (or run .\Invoke-FileNameTransformation.exe)" -ForegroundColor Cyan
Write-Host ""
Write-Host "2. In the GUI:" -ForegroundColor White
Write-Host "   - Tab 6 (Custom Rules): View or add custom regex patterns." -ForegroundColor Gray
Write-Host "   - Tab 1 (File Selection): Browse to folder '$TargetFolderName' and click 'Run Token Analysis'." -ForegroundColor Gray
Write-Host "   - Tab 2 (Transformation): Rearrange or reformat tokens (e.g. Custom:DeptCode, Custom:InvoiceNo)." -ForegroundColor Gray
Write-Host "   - Tab 3 (Preview & Execute): Review transformation and rename files!" -ForegroundColor Gray
Write-Host "============================================================" -ForegroundColor Yellow
Write-Host ""
Write-Host "[v] Demo environment setup complete!" -ForegroundColor Green
