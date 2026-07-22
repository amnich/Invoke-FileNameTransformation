<#
.SYNOPSIS
    Launches the File Name Transformer GUI for bulk renaming and copying.

.DESCRIPTION
    FileNameTransformer.GUI is a WPF-based PowerShell script that enables bulk renaming and copying of files.
    It parses source filenames using structural patterns, applies value mappings from CSV dictionaries,
    and performs text transformations such as casing changes, substrings, replacements, padding, and date formatting.
    The script then builds target filenames from user-defined templates and offers a preview grid to validate
    results before copying or moving files.

    Tokenizer regex: the expression must match every filename segment. Use the named group
    (?<sep>...) for literal separators; all other matches are treated as value fields. The default
    (?<value>[^_\-\s]+)|(?<sep>[_\-\s]+) splits values on underscores, hyphens, and whitespace.
    A second semantic pass recognizes integers, invariant decimals, exact dates/times, GUIDs,
    versions, and configured custom regex types. Recognized composite values can span lexical
    separators. Ambiguous values retain all candidates and require a field type selection.

    Features:
    - Multi-language UI (Polish, English, German).
    - Saveable and reusable profiles in JSON format.
    - Pattern-based source file parsing.
    - Typed field inference with explicit ambiguity resolution.
    - Strict structural matching when applying a selected pattern to other files.
    - External CSV-based mapping support.
    - Preview and validation for collisions, missing values, and invalid characters.
    - Copy or move execution mode with audit logging.
    - Independent options for scanning subfolders and preserving their structure at the destination.

    The saved language in config.json takes precedence over the operating system language. Each normal
    script invocation starts the GUI in an isolated STA PowerShell host so it can be reopened safely
    from the same interactive PowerShell session.

.EXAMPLE
    .\Invoke-FileNameTransformation.ps1
    Launches the application GUI from the current folder.

.EXAMPLE
    powershell.exe -NoProfile -STA -File ".\Invoke-FileNameTransformation.ps1"
    Starts the application from a command prompt, shortcut, or automation scenario. The script starts
    its GUI in an isolated STA child host.

.PARAMETER IsolatedHost
    Internal switch used by the launcher for the isolated GUI process. Do not specify it during normal use.

.NOTES
    Requires Windows PowerShell 5.1 and FileNameTransformation.Core.psm1 beside the development script.
    The development script reads config.json from its own directory. Profiles and logs are stored under
    the current user's AppData folder, with fallback to the script directory or temporary folder when needed.
    When subfolder scanning is enabled, destination folder preservation is controlled separately and is enabled by default.
    The Compliance tab is available only when the Windows USERDOMAIN environment variable contains BGH.
#>

#requires -Version 5.1

[CmdletBinding()]
param(
    [switch]$IsolatedHost
)

if (-not $IsolatedHost) {
    $powershellPath = Join-Path $PSHOME 'powershell.exe'
    if (-not (Test-Path -LiteralPath $powershellPath -PathType Leaf)) {
        $powershellPath = 'powershell.exe'
    }

    & $powershellPath -NoProfile -STA -File $PSCommandPath -IsolatedHost
    return
}

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms
Add-Type -AssemblyName Microsoft.VisualBasic

if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    Write-Warning 'Run in STA mode!'
}
# Always resolve the script's own directory fresh on every run.
# $script:AppRoot is intentionally overwritten later (line ~889) to the AppData
# folder for profiles/logs, so we keep the source location in $script:ScriptRoot.
$script:ScriptRoot = if ([string]::IsNullOrWhiteSpace($PSScriptRoot)) { (Get-Location).Path } else { $PSScriptRoot }
$script:ConfigPath = Join-Path $script:ScriptRoot 'config.json'
$script:CurrentLanguage = $null
$script:Config = [pscustomobject][ordered]@{
    Version         = 2
    Language        = $null
    CustomTypeRules = @()
}
$coreModulePath = Join-Path $script:ScriptRoot 'FileNameTransformation.Core.psm1'
if (-not (Test-Path -LiteralPath $coreModulePath -PathType Leaf)) {
    throw "Missing core module: $coreModulePath"
}
Import-Module $coreModulePath -Force -DisableNameChecking

if (Test-Path $script:ConfigPath) {
    try {
        Write-Verbose $script:ConfigPath
        $config = Get-Content -LiteralPath $script:ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
        $script:Config = ConvertTo-FNTConfig $config
        if ($script:Config.Language -in @('PL', 'EN', 'DE')) {
            $script:CurrentLanguage = $script:Config.Language
            Write-Information "Using saved language: $script:CurrentLanguage"
        }
    }
    catch {
        Write-Warning "Failed to load config file '$script:ConfigPath': $($_.Exception.Message)"
    }
}
$customRuleValidation = Test-FNTCustomTypeRules @($script:Config.CustomTypeRules)
$script:CustomTypeRules = @($customRuleValidation.ValidRules)
foreach ($ruleError in @($customRuleValidation.Errors)) {
    Write-Warning "Custom type rule skipped: $($ruleError.Message)"
}

# Load translations
. (Join-Path $script:ScriptRoot 'src\Translations.ps1')

# Use the OS language only if no valid setting was loaded from config.json.
if ($script:CurrentLanguage -notin @('PL', 'EN', 'DE')) {
    $osLang = (Get-Culture).TwoLetterISOLanguageName.ToUpper()
    $script:CurrentLanguage = if ($osLang -in @('PL', 'EN', 'DE')) { $osLang } else { 'EN' }
}

function T([string]$Key) {
    if ($script:Translations[$script:CurrentLanguage].ContainsKey($Key)) {
        return $script:Translations[$script:CurrentLanguage][$Key]
    }
    return $Key
}

function Throw-StartupFailure([string]$Stage, [Exception]$Exception) {
    $details = "Startup failed during $Stage. Language: $script:CurrentLanguage."
    if ($Exception.InnerException) {
        $details += " Inner exception: $($Exception.InnerException.GetType().FullName): $($Exception.InnerException.Message)"
    }
    throw [System.InvalidOperationException]::new("$details Exception: $($Exception.GetType().FullName): $($Exception.Message)", $Exception)
}

function Get-XamlLoadDiagnostic([string]$Xaml) {
    $results = New-Object System.Collections.Generic.List[string]
    try {
        $testDocument = New-Object System.Xml.XmlDocument
        $testDocument.LoadXml($Xaml)
        $namespaceManager = New-Object System.Xml.XmlNamespaceManager $testDocument.NameTable
        $namespaceManager.AddNamespace('p', 'http://schemas.microsoft.com/winfx/2006/xaml/presentation')
        $tabItems = @($testDocument.SelectNodes('//p:TabItem', $namespaceManager))

        for ($index = 0; $index -lt $tabItems.Count; $index++) {
            $reducedDocument = New-Object System.Xml.XmlDocument
            $reducedDocument.LoadXml($Xaml)
            $reducedNamespaceManager = New-Object System.Xml.XmlNamespaceManager $reducedDocument.NameTable
            $reducedNamespaceManager.AddNamespace('p', 'http://schemas.microsoft.com/winfx/2006/xaml/presentation')
            $tabToRemove = @($reducedDocument.SelectNodes('//p:TabItem', $reducedNamespaceManager))[$index]
            $header = $tabToRemove.GetAttribute('Header')
            [void]$tabToRemove.ParentNode.RemoveChild($tabToRemove)

            try {
                $testReader = New-Object System.Xml.XmlNodeReader $reducedDocument
                [Windows.Markup.XamlReader]::Load($testReader) | Out-Null
                $results.Add("WPF load succeeds when TabItem #$($index + 1) ('$header') is removed.")
            }
            catch {
                $results.Add("WPF load still fails when TabItem #$($index + 1) ('$header') is removed: $($_.Exception.InnerException.Message)")
            }
        }
    }
    catch {
        $results.Add("Unable to run XAML structural probe: $($_.Exception.Message)")
    }

    return ($results -join ' ')
}
#endregion
#region Application Directories and State
$baseAppRoot = if ($env:APPDATA) { Join-Path $env:APPDATA 'FileNameTransformer' } elseif ($script:AppRoot) { $script:AppRoot } else { (Get-Location).Path }
if ([string]::IsNullOrWhiteSpace($baseAppRoot)) {
    $baseAppRoot = [System.IO.Path]::GetTempPath()
}
$script:AppRoot = $baseAppRoot
$script:ProfileRoot = Join-Path $script:AppRoot 'Profiles'
$script:LogRoot = Join-Path $script:AppRoot 'Logs'
New-Item -ItemType Directory -Path $script:ProfileRoot, $script:LogRoot -Force | Out-Null
$script:LogPath = Join-Path $script:LogRoot ('FileNameTransformer_{0:yyyyMMdd_HHmmss}.log' -f (Get-Date))

$script:Patterns = @()
$script:Fields = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
$script:Mappings = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
$script:OutputParts = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
$script:PreviewRows = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
$script:CurrentProfileName = ''
$script:CurrentPattern = $null
#endregion

#region Screen Size
$script:ScreenWidth = [System.Windows.SystemParameters]::PrimaryScreenWidth * 0.9
$script:ScreenHeight = [System.Windows.SystemParameters]::PrimaryScreenHeight * 0.9
#endregion

#region XAML Definition
# Read MainWindow.xaml from disk
$xamlPath = Join-Path $script:ScriptRoot 'MainWindow.xaml'
try {
    if (-not (Test-Path -LiteralPath $xamlPath -PathType Leaf)) {
        throw "Missing UI template: $xamlPath"
    }
    $xamlTemplate = [System.IO.File]::ReadAllText($xamlPath)
}
catch {
    Throw-StartupFailure "reading UI template '$xamlPath'" $_.Exception
}
#endregion

#region Window Creation and Element Binding
try {
    $xamlTemplate = $xamlTemplate.Replace('{ScreenWidth}', [string][int]$script:ScreenWidth)
    $xamlTemplate = $xamlTemplate.Replace('{ScreenHeight}', [string][int]$script:ScreenHeight)
    foreach ($key in $script:Translations[$script:CurrentLanguage].Keys) {
        $translation = $script:Translations[$script:CurrentLanguage][$key]
        if ($null -eq $translation) {
            throw "Translation '$key' has no value."
        }
        $xamlTemplate = $xamlTemplate.Replace("{t:$key}", [string]$translation)
    }
    $script:IsBGHDomainUser = [string]$env:USERDOMAIN -match 'BGH'
    if (-not $script:IsBGHDomainUser) {
        $xamlTemplate = [regex]::Replace(
            $xamlTemplate,
            '(?s)\s*<!-- BGH_COMPLIANCE_START -->.*?<!-- BGH_COMPLIANCE_END -->',
            ''
        )
    }
}
catch {
    Throw-StartupFailure 'applying translations to the UI template' $_.Exception
}

try {
    $unresolvedTranslationTokens = @([regex]::Matches($xamlTemplate, '\{t:[^}]+\}') | ForEach-Object Value | Sort-Object -Unique)
    if ($unresolvedTranslationTokens.Count -gt 0) {
        throw "Missing translations: $($unresolvedTranslationTokens -join ', ')"
    }
    [xml]$xamlDoc = $xamlTemplate
}
catch {
    Throw-StartupFailure "validating translated UI template '$xamlPath'" $_.Exception
}

try {
    $window = [Windows.Markup.XamlReader]::Parse($xamlTemplate)
}
catch {
    $loadDiagnostic = Get-XamlLoadDiagnostic $xamlTemplate
    $loadException = [System.InvalidOperationException]::new("$($_.Exception.Message) XAML structural probe: $loadDiagnostic", $_.Exception)
    Throw-StartupFailure "creating WPF controls from '$xamlPath'" $loadException
}

try {
    $ns = New-Object System.Xml.XmlNamespaceManager $xamlDoc.NameTable
    $ns.AddNamespace('x', 'http://schemas.microsoft.com/winfx/2006/xaml')
    $xamlDoc.SelectNodes('//*[@x:Name]', $ns) | ForEach-Object {
        $name = $_.GetAttribute('Name', 'http://schemas.microsoft.com/winfx/2006/xaml')
        $control = $window.FindName($name)
        if ($null -eq $control) {
            throw "XAML control '$name' was not created."
        }
        Set-Variable -Name $name -Value $control -Scope Script
    }
}
catch {
    Throw-StartupFailure 'binding named WPF controls' $_.Exception
}
#endregion

#region Utility Functions
# Load source modules
. (Join-Path $script:ScriptRoot 'src\UI.ps1')
. (Join-Path $script:ScriptRoot 'src\Analysis.ps1')
. (Join-Path $script:ScriptRoot 'src\Transforms.ps1')
. (Join-Path $script:ScriptRoot 'src\Mappings.ps1')
. (Join-Path $script:ScriptRoot 'src\Preview.ps1')
. (Join-Path $script:ScriptRoot 'src\Profiles.ps1')
. (Join-Path $script:ScriptRoot 'src\Compliance.ps1')

function Log([string]$message, [string]$level = 'INFO') {
    if ([string]::IsNullOrWhiteSpace($script:LogPath)) {
        if ([string]::IsNullOrWhiteSpace($script:LogRoot)) {
            $script:LogRoot = Join-Path $script:AppRoot 'Logs'
        }
        if ([string]::IsNullOrWhiteSpace($script:AppRoot)) {
            $script:AppRoot = [System.IO.Path]::GetTempPath()
        }
        New-Item -ItemType Directory -Path $script:LogRoot -Force | Out-Null
        $script:LogPath = Join-Path $script:LogRoot ('FileNameTransformer_{0:yyyyMMdd_HHmmss}.log' -f (Get-Date))
    }

    $line = '{0:yyyy-MM-dd HH:mm:ss} [{1}] {2}' -f (Get-Date), $level, $message
    Add-Content -LiteralPath $script:LogPath -Encoding UTF8 -Value $line
}

function SetStatus([string]$message) {
    $StatusText.Text = $message
    Log $message
}

function GetLocalizedErrorMessage([Exception]$exception) {
    if ($null -eq $exception -or -not $exception.Data.Contains('FNTCode')) {
        return $exception.Message
    }

    $kindLabel = {
        param($kind)
        if ($kind -eq 'separator') { return (T 'Diag_KindSeparator') }
        return (T 'Diag_KindValue')
    }
    switch ([string]$exception.Data['FNTCode']) {
        'Tokenizer.InvalidRegex' {
            return ((T 'Diag_InvalidRegex') -f $exception.Data['Reason'])
        }
        'Tokenizer.MissingSeparatorGroup' { return (T 'Diag_MissingSepGroup') }
        'Tokenizer.ZeroLength' {
            return ((T 'Diag_ZeroLength') -f $exception.Data['Position'])
        }
        'Tokenizer.IncompleteCoverage' {
            return ((T 'Diag_IncompleteCoverage') -f $exception.Data['Position'])
        }
        'Pattern.TokenCount' {
            return ((T 'Diag_TokenCount') -f $exception.Data['Name'], $exception.Data['Expected'], $exception.Data['Actual'])
        }
        'Pattern.TokenKind' {
            return ((T 'Diag_TokenKind') -f $exception.Data['Name'], $exception.Data['Token'], $exception.Data['Offset'],
                (& $kindLabel $exception.Data['Expected']), (& $kindLabel $exception.Data['Actual']), $exception.Data['Value'])
        }
        'Pattern.Separator' {
            return ((T 'Diag_Separator') -f $exception.Data['Name'], $exception.Data['Token'], $exception.Data['Offset'],
                $exception.Data['Expected'], $exception.Data['Actual'])
        }
        'Pattern.Type' {
            $formatText = if ($exception.Data['Format']) { " ($($exception.Data['Format']))" } else { '' }
            return ((T 'Diag_Type') -f $exception.Data['Name'], $exception.Data['Token'], $exception.Data['Offset'],
                $exception.Data['Value'], $exception.Data['TypeId'], $formatText)
        }
        default { return $exception.Message }
    }
}

function ErrorBox([string]$title, $err) {
    $localizedMessage = GetLocalizedErrorMessage $err.Exception
    $msg = "Komunikat: $localizedMessage"
    $msg += "`nLinia:     $($err.InvocationInfo.ScriptLineNumber)"
    $msg += "`nLog:       $script:LogPath"
    Log "$title | $localizedMessage" 'ERROR'
    [Windows.MessageBox]::Show($msg, $title, 'OK', 'Error') | Out-Null
}

function FolderDialog([string]$caption, [string]$initial) {
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = $caption

    $initialPath = [string]$initial
    if (-not [string]::IsNullOrWhiteSpace($initialPath) -and (Test-Path -LiteralPath $initialPath -PathType Container)) {
        $dlg.SelectedPath = $initialPath
    }

    if ($dlg.ShowDialog() -eq 'OK') { $dlg.SelectedPath }
}

function FileDialog {
    $dlg = New-Object Microsoft.Win32.OpenFileDialog
    $dlg.Filter = 'CSV (*.csv)|*.csv|Tekst (*.txt)|*.txt|Wszystkie (*.*)|*.*'
    if ($dlg.ShowDialog()) { $dlg.FileName }
}

function UpdateUI {
    # Force WPF dispatcher to process pending UI updates
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
        [Action] {},
        [System.Windows.Threading.DispatcherPriority]::Background
    )
}

#endregion

#region Event Handlers
# Load event handlers
. (Join-Path $script:ScriptRoot 'src\Events.ps1')

# --- Folder browse ---
$BrowseSource.Add_Click({
        $p = FolderDialog (T 'Msg_SelectSource') $SourcePath.Text
        if ($p) { $SourcePath.Text = $p }
    })
$BrowseDestination.Add_Click({
        $p = FolderDialog (T 'Msg_SelectDest') $DestinationPath.Text
        if ($p) { $DestinationPath.Text = $p }
    })

# --- Tab 2: Analysis ---
$Analyze.Add_Click({
        try { AnalyzePatterns; SetStatus (T 'Status_AnalysisDone') }
        catch { ErrorBox (T 'Err_Analysis') $_ }
    })
$ExtensionFilter.Add_SelectionChanged({
        try { BuildPatternList } catch {}
    })
$PatternGrid.Add_SelectionChanged({
        if ($PatternGrid.SelectedItem) {
            SetPattern $PatternGrid.SelectedItem
        }
    })

# --- Tab 3: Fields ---
$FieldGrid.Add_SelectionChanged({
        $f = $FieldGrid.SelectedItem
        if ($f) {
            $FieldName.Text = $f.Name
            # Select matching role item
            $FieldRole.SelectedItem = @($FieldRole.Items | Where-Object { $_.Content -eq $f.Role })[0]
            if (-not $FieldRole.SelectedItem) { $FieldRole.SelectedIndex = 0 }
            $typeOptions = @(GetFieldTypeOptions $f)
            $FieldType.ItemsSource = $typeOptions
            $selectedType = @($typeOptions | Where-Object {
                    $_.Id -eq $f.SelectedTypeId -and [string]$_.Format -eq [string]$f.SelectedFormat
                } | Select-Object -First 1)
            if ($selectedType.Count -gt 0) { $FieldType.SelectedItem = $selectedType[0] }
            else { $FieldType.SelectedIndex = 0 }
            $CandidateInfo.Text = GetFieldCandidateSummary $f
            # Show transforms for this field
            $TransformList.ItemsSource = $null
            if ($f.Transforms) {
                $TransformList.ItemsSource = @($f.Transforms)
            }
        }
    })

$FieldApply.Add_Click({
        try {
            $f = $FieldGrid.SelectedItem
            if (-not $f) { throw (T 'Err_SelFieldTab') }
            $newName = $FieldName.Text.Trim()
            if (-not $newName) { throw (T 'Err_FieldName') }
            $oldName = $f.Name

            $f.Name = $newName
            $f.Role = [string]$FieldRole.SelectedItem.Content
            $selectedType = $FieldType.SelectedItem
            $f.SelectedTypeId = if ($selectedType) { [string]$selectedType.Id } else { 'Auto' }
            $f.SelectedFormat = if ($selectedType) { [string]$selectedType.Format } else { $null }
            $f.EffectiveType = GetEffectiveTypeLabel $f
            $f.TypeStatus = GetFieldTypeStatus $f
            $f.CandidateSummary = GetFieldCandidateSummary $f

            # Update references in OutputParts and Mappings when name changed
            if ($oldName -ne $newName) {
                foreach ($p in $script:OutputParts) {
                    if ($p.Type -eq 'Field' -and $p.Value -eq $oldName) {
                        $p.Value = $newName
                        $p.Display = "$(T 'Tag_Field') $newName"
                    }
                }
                foreach ($m in $script:Mappings) {
                    $changed = $false
                    if ($m.InputField -eq $oldName) { $m.InputField = $newName; $changed = $true }
                    if ($m.OutputField -eq $oldName) { $m.OutputField = $newName; $changed = $true }
                    if ($changed) { $m.Display = "$($m.Name): $($m.InputField) → $($m.OutputField)" }
                }
                $OutputList.ItemsSource = $null
                $OutputList.ItemsSource = $script:OutputParts
                $MappingList.ItemsSource = $null
                $MappingList.ItemsSource = $script:Mappings
            }

            $FieldGrid.ItemsSource = $null
            $FieldGrid.ItemsSource = $script:Fields
            RefreshFieldSelector
            UpdateOutputExample
        }
        catch { ErrorBox (T 'Err_Field') $_ }
    })

# --- Tab 3: Transforms ---
$TransformAdd.Add_Click({
        try {
            $f = $FieldGrid.SelectedItem
            if (-not $f) { throw (T 'Err_SelFieldTrans') }
            $result = ShowTransformDialog $f
            if ($result) {
                if (-not $f.Transforms) {
                    $f.Transforms = [System.Collections.ArrayList]::new()
                }
                [void]$f.Transforms.Add($result)
                $TransformList.ItemsSource = $null
                $TransformList.ItemsSource = @($f.Transforms)
                UpdateOutputExample
            }
        }
        catch { ErrorBox (T 'Err_Transform') $_ }
    })

$TransformRemove.Add_Click({
        $f = $FieldGrid.SelectedItem
        $i = $TransformList.SelectedIndex
        if ($f -and $f.Transforms -and $i -ge 0) {
            $f.Transforms.RemoveAt($i)
            $TransformList.ItemsSource = $null
            $TransformList.ItemsSource = @($f.Transforms)
            UpdateOutputExample
        }
    })

# --- Tab 3: Mappings ---
$MappingAdd.Add_Click({
        try { AddMappingDialog; UpdateOutputExample }
        catch { ErrorBox (T 'Err_Mapping') $_ }
    })

$MappingEdit.Add_Click({
        $i = $MappingList.SelectedIndex
        if ($i -lt 0) { throw (T 'Err_SelMapping') }
        try { AddMappingDialog $script:Mappings[$i]; UpdateOutputExample }
        catch { ErrorBox (T 'Err_Mapping') $_ }
    })

$MappingRemove.Add_Click({
        $i = $MappingList.SelectedIndex
        if ($i -ge 0) {
            # Remove virtual field if no other mapping produces it
            $removedOutput = $script:Mappings[$i].OutputField
            $script:Mappings.RemoveAt($i)

            $stillUsed = $false
            foreach ($m in $script:Mappings) {
                if ($m.OutputField -eq $removedOutput) { $stillUsed = $true; break }
            }
            if (-not $stillUsed) {
                $toRemove = $null
                foreach ($f in $script:Fields) {
                    if ($f.IsVirtual -and $f.Name -eq $removedOutput) { $toRemove = $f; break }
                }
                if ($toRemove) { $script:Fields.Remove($toRemove) }
            }

            $MappingList.ItemsSource = $null
            $MappingList.ItemsSource = $script:Mappings
            $FieldGrid.ItemsSource = $null
            $FieldGrid.ItemsSource = $script:Fields
            RefreshFieldSelector
            UpdateOutputExample
        }
    })

$MappingUp.Add_Click({
        $i = $MappingList.SelectedIndex
        if ($i -gt 0) {
            $x = $script:Mappings[$i]
            $script:Mappings.RemoveAt($i)
            $script:Mappings.Insert($i - 1, $x)
            $MappingList.ItemsSource = $null
            $MappingList.ItemsSource = $script:Mappings
            $MappingList.SelectedIndex = $i - 1
        }
    })

$MappingDown.Add_Click({
        $i = $MappingList.SelectedIndex
        if ($i -ge 0 -and $i -lt $script:Mappings.Count - 1) {
            $x = $script:Mappings[$i]
            $script:Mappings.RemoveAt($i)
            $script:Mappings.Insert($i + 1, $x)
            $MappingList.ItemsSource = $null
            $MappingList.ItemsSource = $script:Mappings
            $MappingList.SelectedIndex = $i + 1
        }
    })

# --- Tab 4: Output name builder ---
$OutputAddField.Add_Click({
        try {
            $fieldName = [string]$FieldSelector.SelectedItem
            if (-not $fieldName) { throw (T 'Err_SelFieldCombo') }
            $script:OutputParts.Add([pscustomobject]@{
                    Type    = 'Field'
                    Value   = $fieldName
                    Display = "$(T 'Tag_Field') $fieldName"
                })
            $OutputList.ItemsSource = $null
            $OutputList.ItemsSource = $script:OutputParts
            UpdateOutputExample
        }
        catch { ErrorBox (T 'Title_Error') $_ }
    })

$OutputAddText.Add_Click({
        $v = $OutputText.Text
        if ($v) {
            $script:OutputParts.Add([pscustomobject]@{
                    Type    = 'Text'
                    Value   = $v
                    Display = "$(T 'Tag_Text') $v"
                })
            $OutputText.Clear()
            $OutputList.ItemsSource = $null
            $OutputList.ItemsSource = $script:OutputParts
            UpdateOutputExample
        }
    })

$OutputAddSeparator.Add_Click({
        $script:OutputParts.Add([pscustomobject]@{
                Type    = 'Text'
                Value   = '_'
                Display = "$(T 'Tag_Separator') _"
            })
        $OutputList.ItemsSource = $null
        $OutputList.ItemsSource = $script:OutputParts
        UpdateOutputExample
    })

$OutputRemove.Add_Click({
        $i = $OutputList.SelectedIndex
        if ($i -ge 0) {
            $script:OutputParts.RemoveAt($i)
            $OutputList.ItemsSource = $null
            $OutputList.ItemsSource = $script:OutputParts
            UpdateOutputExample
        }
    })

$OutputUp.Add_Click({
        $i = $OutputList.SelectedIndex
        if ($i -gt 0) {
            $x = $script:OutputParts[$i]
            $script:OutputParts.RemoveAt($i)
            $script:OutputParts.Insert($i - 1, $x)
            $OutputList.ItemsSource = $null
            $OutputList.ItemsSource = $script:OutputParts
            $OutputList.SelectedIndex = $i - 1
            UpdateOutputExample
        }
    })

$OutputDown.Add_Click({
        $i = $OutputList.SelectedIndex
        if ($i -ge 0 -and $i -lt $script:OutputParts.Count - 1) {
            $x = $script:OutputParts[$i]
            $script:OutputParts.RemoveAt($i)
            $script:OutputParts.Insert($i + 1, $x)
            $OutputList.ItemsSource = $null
            $OutputList.ItemsSource = $script:OutputParts
            $OutputList.SelectedIndex = $i + 1
            UpdateOutputExample
        }
    })

$KeepExtension.Add_Checked({ $NewExtension.IsEnabled = $false })
$KeepExtension.Add_Unchecked({ $NewExtension.IsEnabled = $true })

# --- Tab 5: Preview and execution ---
$BuildPreview.Add_Click({
        try { FullBuildPreview; SetStatus (T 'Status_PrevBuilt') }
        catch { ErrorBox (T 'Err_Preview') $_ }
    })

$PreviewFilter.Add_SelectionChanged({ RefreshPreviewGrid })

$ExportAudit.Add_Click({
        try {
            if (-not $script:PreviewRows.Count) { throw (T 'Err_BuildPrev1') }
            $dlg = New-Object Microsoft.Win32.SaveFileDialog
            $dlg.Filter = 'CSV (*.csv)|*.csv'
            $dlg.FileName = 'FileNameTransformationAudit.csv'
            if ($dlg.ShowDialog()) {
                $script:PreviewRows | Export-Csv -LiteralPath $dlg.FileName -NoTypeInformation -Encoding UTF8 -UseCulture
                SetStatus "$(T 'Status_Exported') $($dlg.FileName)"
            }
        }
        catch { ErrorBox (T 'Err_Export') $_ }
    })

$OpenLog.Add_Click({
        Start-Process notepad.exe "`"$script:LogPath`""
    })

$Execute.Add_Click({
        try { ExecuteCopy }
        catch { ErrorBox (T 'Err_Execute') $_ }
    })

# --- Tab 1: Profiles ---
$ProfileNew.Add_Click({
        $script:Fields.Clear()
        $script:Mappings.Clear()
        $script:OutputParts.Clear()
        $FieldGrid.ItemsSource = $script:Fields
        $MappingList.ItemsSource = $script:Mappings
        $OutputList.ItemsSource = $script:OutputParts
        $TransformList.ItemsSource = $null
        RefreshFieldSelector
        $script:CurrentProfileName = ''
        $CurrentProfile.Text = (T 'Txt_Unsaved')
        UpdateOutputExample
    })

$ProfileSave.Add_Click({
        try { SaveProfile }
        catch { ErrorBox (T 'Err_SaveProfile') $_ }
    })

$ProfileLoad.Add_Click({
        try {
            if (-not $ProfileList.SelectedItem) { throw (T 'Err_SelProfList') }
            LoadProfile $ProfileList.SelectedItem.Path
        }
        catch { ErrorBox (T 'Err_LoadProfile') $_ }
    })

$ProfileCopy.Add_Click({
        try {
            if (-not $ProfileList.SelectedItem) { throw (T 'Err_SelProfCopy') }
            LoadProfile $ProfileList.SelectedItem.Path
            $script:CurrentProfileName = (T 'Prefix_Copy') + $script:CurrentProfileName
            SaveProfile
        }
        catch { ErrorBox (T 'Err_CopyProfile') $_ }
    })

$ProfileDelete.Add_Click({
        try {
            if (-not $ProfileList.SelectedItem) { throw (T 'Err_SelProfDel') }
            $name = $ProfileList.SelectedItem.Name
            $confirm = [Windows.MessageBox]::Show(
                "$(T 'Msg_DelProf') '$name'?", (T 'Title_Confirm'), 'YesNo', 'Warning'
            )
            if ($confirm -eq 'Yes') {
                Remove-Item $ProfileList.SelectedItem.Path -Force
                RefreshProfiles
                SetStatus "$(T 'Status_ProfDel') $name"
            }
        }
        catch { ErrorBox (T 'Err_DelProfile') $_ }
    })

#endregion

#region Initialization
$LanguageSelector.Add_SelectionChanged({
        if ($LanguageSelector.SelectedItem) {
            $tag = $LanguageSelector.SelectedItem.Tag
            if ($tag -and $tag -ne $script:CurrentLanguage) {
                $script:Config = Set-FNTConfigLanguage -Config $script:Config -Language $tag
                $script:Config | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $script:ConfigPath -Encoding UTF8
                [Windows.MessageBox]::Show((T 'Msg_ConfirmRestart'), (T 'Title_Info'), 'OK', 'Information') | Out-Null
            }
        }
    })
$LogText.Text = "Log: $script:LogPath"
$FieldRole.SelectedIndex = 0
RefreshProfiles
UpdateOutputExample

$window.Add_Closed({ Log (T 'Log_AppClosed') })
SetStatus (T 'Status_Ready')

[void]$window.ShowDialog()
#endregion