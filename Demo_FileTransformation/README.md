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
