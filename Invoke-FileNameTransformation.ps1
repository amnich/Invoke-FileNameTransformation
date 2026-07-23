<#
.SYNOPSIS
    Launches the File Name Transformer WPF GUI application for bulk file renaming, structure transformation, and compliance auditing.

.DESCRIPTION
    Invoke-FileNameTransformation is a feature-rich, multi-language WPF application built for Windows PowerShell 5.1.
    It enables advanced pattern-based parsing of source filenames, value transformations, CSV dictionary lookups,
    metadata extraction, and live validation preview before performing copy or move operations.

    Key Features:
    - 7 Interactive GUI Tabs: Profile & Folder Setup, Structural Analysis, Field & Mapping Configuration, Destination Template Builder, Live Preview & Execution, Live Custom Type Regex Manager & Tester, and Name Compliance Scanner.
    - Two-Pass Tokenizer Engine: Combines lexical regex splitting with semantic composite value merging for dates (exact valid calendar formats), invariant decimals, integers, GUIDs, versions (2 to 4 components), and custom regex types.
    - Strict Structural Enforcement: Validates token counts, separator literals, and semantic field data types across file batches, flagging structural mismatches with precise error diagnostics.
    - External CSV/TXT Mappings: Connects lookup files to map extracted values to virtual fields for enriched destination filenames.
    - Extensive Field Transformation Suite: Casing adjustments (UPPERCASE, lowercase, Title Case), text trimming (Substring), padding, replacements, date formatting, math operations, diacritic transliteration, regex replace, PowerShell script expressions, and zero-padded sequential counters.
    - Metadata Extraction: Reads EXIF photo tags, Office OpenXML properties, COM Shell audio tags, filesystem timestamps, and computes MD5/SHA256 content hashes.
    - Process Isolation & Reentrancy: Automatically spawns an isolated STA child PowerShell process so the GUI can be reopened repeatedly from an interactive console without WPF apartment or event handler state conflicts.
    - Saveable Profiles & Config: Saves complex renamer setups as JSON profiles (Version 2 schema). Persists language preferences (EN, PL, DE) and dark/light themes in script-local config.json.

.EXAMPLE
    .\Invoke-FileNameTransformation.ps1
    Launches the application GUI using the default STA launcher.

.EXAMPLE
    powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File ".\Invoke-FileNameTransformation.ps1"
    Launches the application from a command prompt, shortcut, or task scheduler. The script launches an isolated STA host process.

.PARAMETER IsolatedHost
    Internal switch parameter passed automatically when launching the isolated GUI host process.
    Do not specify this switch manually during normal invocation.

.INPUTS
    None. Interactive GUI application.

.OUTPUTS
    None. Performs file copy/move operations, updates profile JSON files, creates audit logs in AppData, and exports CSV reports.

.NOTES
    Prerequisites:
    - Windows PowerShell 5.1 on Windows OS with WPF/WinForms support.
    - Requires FileNameTransformation.Core.psm1 and src/ module files in the application directory.
    - Reads config.json from the script directory. Profile JSON files and timestamped log files are saved under %APPDATA%\FileNameTransformer with automatic fallback to script folder or temp directory.
    - Tab 0 (Name Compliance Scanner) is activated when the USERDOMAIN environment variable contains 'BGH'.
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
$script:CurrentTheme = 'Dark'
$script:Config = [pscustomobject][ordered]@{
    Version         = 2
    Language        = $null
    Theme           = 'Dark'
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
        if ($script:Config.Theme -in @('Light', 'Dark')) {
            $script:CurrentTheme = $script:Config.Theme
            Write-Information "Using saved theme: $script:CurrentTheme"
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
$script:MetadataCache = @{}
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
        $escapedTranslation = [System.Security.SecurityElement]::Escape([string]$translation)
        $xamlTemplate = $xamlTemplate.Replace("{t:$key}", $escapedTranslation)
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
        if ($null -ne $control) {
            Set-Variable -Name $name -Value $control -Scope Script
        }
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
    $msg = "$(T 'Message'): $localizedMessage"
    $msg += "`n$(T 'Line'):     $($err.InvocationInfo.ScriptLineNumber)"
    $msg += "`n$(T 'Log'):       $script:LogPath"
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

#endregion

#region Initialization
Set-WPFWindowTheme $window $script:CurrentTheme
$window.Add_SourceInitialized({
        Set-WPFWindowTheme $window $script:CurrentTheme
    })

if ($LanguageSelector) {
    $matchedLangItem = @($LanguageSelector.Items | Where-Object { $_.Tag -eq $script:CurrentLanguage })[0]
    if ($matchedLangItem) { $LanguageSelector.SelectedItem = $matchedLangItem }
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
}

if ($ThemeSelector) {
    $matchedThemeItem = @($ThemeSelector.Items | Where-Object { $_.Tag -eq $script:CurrentTheme })[0]
    if ($matchedThemeItem) { $ThemeSelector.SelectedItem = $matchedThemeItem }
    $ThemeSelector.Add_SelectionChanged({
            if ($ThemeSelector.SelectedItem) {
                $tag = $ThemeSelector.SelectedItem.Tag
                if ($tag -and $tag -ne $script:CurrentTheme) {
                    $script:CurrentTheme = $tag
                    $script:Config = Set-FNTConfigTheme -Config $script:Config -Theme $tag
                    $script:Config | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $script:ConfigPath -Encoding UTF8
                    Set-WPFWindowTheme $window $script:CurrentTheme
                }
            }
        })
}

$LogText.Text = "Log: $script:LogPath"
$FieldRole.SelectedIndex = 0
RefreshProfiles
UpdateOutputExample

$window.Add_Closed({ Log (T 'Log_AppClosed') })
SetStatus (T 'Status_Ready')

[void]$window.ShowDialog()
#endregion