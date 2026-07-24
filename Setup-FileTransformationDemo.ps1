<#
.SYNOPSIS
    Creates sample files and mapping data for the File Name Transformer mapping demo.

.DESCRIPTION
    Creates 20 PDF placeholder files with the input structure
    EmployeeID_yyyyMMdd_CourseID_Company.pdf, two CSV mapping files, and a
    Markdown walkthrough for transforming the files to
    MM_yyyy_DisplayName_Company_CourseName.pdf.

.PARAMETER TargetFolderName
    Name or path of the demo folder to create. Defaults to Demo_FileTransformation.

.EXAMPLE
    .\Setup-FileTransformationDemo.ps1
#>

[CmdletBinding()]
param(
    [string]$TargetFolderName = 'Demo_FileTransformation'
)

$ErrorActionPreference = 'Stop'
$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Definition
$demoFolderPath = Join-Path $scriptDirectory $TargetFolderName

if (-not (Test-Path -LiteralPath $demoFolderPath -PathType Container)) {
    New-Item -ItemType Directory -Path $demoFolderPath | Out-Null
    Write-Host "Created demo directory: $demoFolderPath" -ForegroundColor Green
}
else {
    Write-Host "Using existing demo directory: $demoFolderPath" -ForegroundColor Yellow
}

$employees = @(
    [pscustomobject]@{ EmployeeID = 'E1001'; DisplayName = 'Kowalski, Jan' }
    [pscustomobject]@{ EmployeeID = 'E1002'; DisplayName = 'Nowak, Anna' }
    [pscustomobject]@{ EmployeeID = 'E1003'; DisplayName = 'Wisniewski, Piotr' }
    [pscustomobject]@{ EmployeeID = 'E1004'; DisplayName = 'Wojcik, Marta' }
    [pscustomobject]@{ EmployeeID = 'E1005'; DisplayName = 'Kaminski, Tomasz' }
    [pscustomobject]@{ EmployeeID = 'E1006'; DisplayName = 'Lewandowska, Ewa' }
    [pscustomobject]@{ EmployeeID = 'E1007'; DisplayName = 'Zielinski, Marek' }
    [pscustomobject]@{ EmployeeID = 'E1008'; DisplayName = 'Szymanska, Olga' }
    [pscustomobject]@{ EmployeeID = 'E1009'; DisplayName = 'Dabrowski, Pawel' }
    [pscustomobject]@{ EmployeeID = 'E1010'; DisplayName = 'Kozlowska, Iwona' }
)

$courses = @(
    [pscustomobject]@{ CourseID = 'C101'; CourseName = 'PowerShell_Basics' }
    [pscustomobject]@{ CourseID = 'C102'; CourseName = 'Excel_Reporting' }
    [pscustomobject]@{ CourseID = 'C103'; CourseName = 'Information_Security' }
    [pscustomobject]@{ CourseID = 'C104'; CourseName = 'Project_Management' }
    [pscustomobject]@{ CourseID = 'C105'; CourseName = 'Windows_Administration' }
)

$employeeMappingPath = Join-Path $demoFolderPath 'EmployeeMapping.csv'
$courseMappingPath = Join-Path $demoFolderPath 'CourseMapping.csv'
$employees | Export-Csv -LiteralPath $employeeMappingPath -NoTypeInformation -Encoding UTF8
$courses | Export-Csv -LiteralPath $courseMappingPath -NoTypeInformation -Encoding UTF8

$dates = @(
    '20260105', '20260112', '20260119', '20260126', '20260202',
    '20260209', '20260216', '20260223', '20260302', '20260309',
    '20260316', '20260323', '20260406', '20260413', '20260420',
    '20260504', '20260511', '20260518', '20260601', '20260608'
)

$sampleFiles = for ($index = 0; $index -lt 20; $index++) {
    $employee = $employees[$index % $employees.Count]
    $course = $courses[$index % $courses.Count]
    '{0}_{1}_{2}_BGH.pdf' -f $employee.EmployeeID, $dates[$index], $course.CourseID
}

foreach ($fileName in $sampleFiles) {
    $filePath = Join-Path $demoFolderPath $fileName
    if (-not (Test-Path -LiteralPath $filePath -PathType Leaf)) {
        Set-Content -LiteralPath $filePath -Value 'File Name Transformer mapping demo placeholder.' -Encoding UTF8
    }
}

$readmePath = Join-Path $demoFolderPath 'README.md'
$readmeContent = @'
# File Transformation and Mapping Demo

This folder contains 20 sample files and two CSV mapping files for replaying a complete File Name Transformer workflow.

## Input and Expected Result

The sample files use this input format:

```text
EmployeeID_yyyyMMdd_CourseID_Company.pdf
E1001_20260105_C101_BGH.pdf
```

After configuration, the result is:

```text
MM_yyyy_DisplayName_Company_CourseName.pdf
01_2026_KowalskiJ_BGH_PowerShell_Basics.pdf
```

`DisplayName` is mapped from `Kowalski, Jan` and then transformed to `KowalskiJ`.

## Demo Files

- `EmployeeMapping.csv`: maps `EmployeeID` to `DisplayName`.
- `CourseMapping.csv`: maps `CourseID` to `CourseName`.
- Twenty placeholder PDF files use the same four-part filename structure.

## Replay Steps

### 1. Start the application

From the application directory, run:

```powershell
.\Invoke-FileNameTransformation.ps1
```

### 2. Analyse the sample files

1. On **File Analysis**, select this `Demo_FileTransformation` folder.
2. Keep the `.pdf` extension filter selected.
3. Select the one detected filename structure. It contains 20 files.
4. On **Fields & Mappings**, rename the detected fields to:

| Input position | Field name | Role | Data type |
| --- | --- | --- | --- |
| `E1001` | `EmployeeID` | Identifier | Auto |
| `20260105` | `TrainingDate` | Date | Date/time (yyyyMMdd) |
| `C101` | `CourseID` | Identifier | Auto |
| `BGH` | `Company` | Constant text | Auto |

Click **Apply Field** after configuring each field.

### 3. Transform the date

1. Select `TrainingDate` in the field table.
2. Click **Add Transformation**.
3. Select **Date Format**.
4. Set input format to `yyyyMMdd`.
5. Set output format to `MM_yyyy`.
6. Click **Add**.

### 4. Add the employee mapping

1. In **Mappings**, click **Add Mapping**.
2. Enter these values:

| Setting | Value |
| --- | --- |
| Mapping name | `Employee ID to Display Name` |
| Input field | `EmployeeID` |
| Output field | `DisplayName` |
| Mapping file | `EmployeeMapping.csv` in this demo folder |
| Key column | `EmployeeID` |
| Value column | `DisplayName` |

3. Click **Add**. `DisplayName` is added as a virtual field.

### 5. Transform the mapped display name

1. Select the virtual `DisplayName` field in the field table.
2. Click **Add Transformation** and select **PowerShell Expression**.
3. Enter this expression exactly:

```powershell
if ($_ -match '^\s*([^,]+),\s*(.)') { "$($matches[1])$($matches[2])" } else { $_ }
```

4. Click **Add**. For example, `Kowalski, Jan` becomes `KowalskiJ`.

### 6. Add the course mapping

1. In **Mappings**, click **Add Mapping**.
2. Enter these values:

| Setting | Value |
| --- | --- |
| Mapping name | `Course ID to Course Name` |
| Input field | `CourseID` |
| Output field | `CourseName` |
| Mapping file | `CourseMapping.csv` in this demo folder |
| Key column | `CourseID` |
| Value column | `CourseName` |

3. Click **Add**. `CourseName` is added as a virtual field.

### 7. Build the destination filename

On **Destination Name**, add fields and separators in exactly this order:

```text
TrainingDate _ DisplayName _ Company _ CourseName
```

Keep **Keep original extension** enabled. The output example should resemble:

```text
01_2026_KowalskiJ_BGH_PowerShell_Basics.pdf
```

### 8. Preview and execute safely

1. Open **Preview & Execution** and Select a Destination folder for the transformed files. It should be different from the source folder.
2. Click **Build Preview**.
3. Confirm that all 20 proposed names follow the target format.
4. Select **Copy** first, choose a separate destination folder, and execute the transformation.
5. When the copied results are correct, repeat with **Move** only if you want to rename the original files.

## Expected Examples

| Source | Result |
| --- | --- |
| `E1001_20260105_C101_BGH.pdf` | `01_2026_KowalskiJ_BGH_PowerShell_Basics.pdf` |
| `E1002_20260112_C102_BGH.pdf` | `01_2026_NowakA_BGH_Excel_Reporting.pdf` |
| `E1003_20260119_C103_BGH.pdf` | `01_2026_WisniewskiP_BGH_Information_Security.pdf` |

The demo uses underscore characters in course names because Windows filenames cannot contain a slash and underscores preserve a readable multi-word course name.
'@

Set-Content -LiteralPath $readmePath -Value $readmeContent -Encoding UTF8

Write-Host "Created $($sampleFiles.Count) sample PDF files." -ForegroundColor Green
Write-Host "Created mapping file: $employeeMappingPath" -ForegroundColor Green
Write-Host "Created mapping file: $courseMappingPath" -ForegroundColor Green
Write-Host "Created walkthrough: $readmePath" -ForegroundColor Green