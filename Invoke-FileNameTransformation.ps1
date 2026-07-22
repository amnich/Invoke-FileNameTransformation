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