# UI.ps1 — Utility functions, logging, dialogs, and field-type helpers.
# Dot-sourced by the main script; operates in $script: scope.

#region Utility Functions

function Log([string]$message, [string]$level = 'INFO') {
    if ([string]::IsNullOrWhiteSpace($script:AppRoot)) {
        $script:AppRoot = [System.IO.Path]::GetTempPath()
    }
    if ([string]::IsNullOrWhiteSpace($script:LogRoot)) {
        $script:LogRoot = Join-Path $script:AppRoot 'Logs'
    }
    if ([string]::IsNullOrWhiteSpace($script:LogPath)) {
        New-Item -ItemType Directory -Path $script:LogRoot -Force | Out-Null
        $script:LogPath = Join-Path $script:LogRoot ('FileNameTransformer_{0:yyyyMMdd_HHmmss}.log' -f (Get-Date))
    }

    $line = '{0:yyyy-MM-dd HH:mm:ss} [{1}] {2}' -f (Get-Date), $level, $message
    Add-Content -LiteralPath $script:LogPath -Encoding UTF8 -Value $line
}

function SetStatus([string]$message) {
    if ($null -ne $StatusText) {
        $StatusText.Text = $message
    }
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
    $lineNum = if ($err.InvocationInfo) { $err.InvocationInfo.ScriptLineNumber } else { 'N/A' }
    $msg = "Komunikat: $localizedMessage"
    $msg += "`nLinia:     $lineNum"
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

function Get-DirectorySuggestions([string]$Path) {
    try {
        $typedPath = if ($Path) { $Path.Trim() } else { '' }
        if ([string]::IsNullOrWhiteSpace($typedPath)) {
            return @(Get-PSDrive -PSProvider FileSystem |
                Where-Object { $_.Root } |
                ForEach-Object { $_.Root } |
                Sort-Object -Unique)
        }

        if ($typedPath -match '^[A-Za-z]:$') {
            $typedPath += '\'
        }

        $hasTrailingSeparator = $typedPath.EndsWith('\') -or $typedPath.EndsWith('/')
        $parentPath = if ($hasTrailingSeparator) { $typedPath } else { Split-Path -Path $typedPath -Parent }
        $leafPrefix = if ($hasTrailingSeparator) { '' } else { Split-Path -Path $typedPath -Leaf }

        if ([string]::IsNullOrWhiteSpace($parentPath)) {
            $parentPath = (Get-Location).Path
        }
        if (-not (Test-Path -LiteralPath $parentPath -PathType Container)) {
            return @()
        }

        return @(Get-ChildItem -LiteralPath $parentPath -Directory -ErrorAction Stop |
            Where-Object { $_.Name.StartsWith($leafPrefix, [StringComparison]::OrdinalIgnoreCase) } |
            Sort-Object Name |
            Select-Object -First 20 |
            ForEach-Object { $_.FullName })
    }
    catch {
        return @()
    }
}

function Update-PathSuggestions($TextBox, $Popup, $SuggestionList) {
    if ($null -eq $TextBox -or $null -eq $Popup -or $null -eq $SuggestionList) { return }
    $suggestions = @(Get-DirectorySuggestions $TextBox.Text)
    $SuggestionList.ItemsSource = $suggestions
    $Popup.IsOpen = $TextBox.IsKeyboardFocusWithin -and $suggestions.Count -gt 0
}

function Apply-PathSuggestion($TextBox, $Popup, $SuggestionList) {
    $selection = [string]$SuggestionList.SelectedItem
    if (-not [string]::IsNullOrWhiteSpace($selection)) {
        $TextBox.Text = $selection
        $TextBox.CaretIndex = $TextBox.Text.Length
    }
    $Popup.IsOpen = $false
    $TextBox.Focus() | Out-Null
}

function FileDialog {
    $dlg = New-Object Microsoft.Win32.OpenFileDialog
    $csvFilter = if (Get-Command 'T' -ErrorAction SilentlyContinue) { T 'Filter_CSV' } else { 'CSV (*.csv)' }
    $txtFilter = if (Get-Command 'T' -ErrorAction SilentlyContinue) { T 'Filter_Text' } else { 'Text (*.txt)' }
    $allFilter = if (Get-Command 'T' -ErrorAction SilentlyContinue) { T 'Filter_AllFiles' } else { 'All Files (*.*)' }
    $dlg.Filter = "$csvFilter|*.csv|$txtFilter|*.txt|$allFilter|*.*"
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

#region Field Type Helpers

function TokenTypeLabel([string]$typeId) {
    switch ($typeId) {
        'Integer' { return (T 'Type_Integer') }
        'Decimal' { return (T 'Type_Decimal') }
        'DateTime' { return (T 'Type_DateTime') }
        'Guid' { return (T 'Type_Guid') }
        'Version' { return (T 'Type_Version') }
        'Ambiguous' { return (T 'Type_Ambiguous') }
        default { return (T 'Type_Text') }
    }
}

function GetFieldTypeOptions($field) {
    $options = @(
        [pscustomobject]@{ Id = 'Auto'; Format = $null; Label = (T 'Type_Auto') }
        [pscustomobject]@{ Id = 'Text'; Format = $null; Label = (T 'Type_Text') }
        [pscustomobject]@{ Id = 'Integer'; Format = $null; Label = (T 'Type_Integer') }
        [pscustomobject]@{ Id = 'Decimal'; Format = $null; Label = (T 'Type_Decimal') }
        [pscustomobject]@{ Id = 'Guid'; Format = $null; Label = (T 'Type_Guid') }
        [pscustomobject]@{ Id = 'Version'; Format = $null; Label = (T 'Type_Version') }
    )

    foreach ($candidate in @($field.CandidateTypes)) {
        if ($candidate.TypeId -eq 'DateTime') {
            $label = "$(T 'Type_DateTime') ($($candidate.Format))"
            $options += [pscustomobject]@{ Id = 'DateTime'; Format = [string]$candidate.Format; Label = $label }
        }
        elseif ([string]$candidate.TypeId -like 'Custom:*') {
            $customLabel = if ($candidate.PSObject.Properties['DisplayName'] -and $candidate.DisplayName) {
                [string]$candidate.DisplayName
            }
            elseif ($candidate.PSObject.Properties['RuleId'] -and $candidate.RuleId) {
                [string]$candidate.RuleId
            }
            else {
                [string]$candidate.TypeId
            }
            $options += [pscustomobject]@{
                Id     = [string]$candidate.TypeId
                Format = $null
                Label  = $customLabel
            }
        }
    }

    $seen = @{}
    @($options | Where-Object {
            $key = "$($_.Id)|$($_.Format)"
            if ($seen.ContainsKey($key)) { return $false }
            $seen[$key] = $true
            return $true
        })
}

function GetEffectiveTypeLabel($field) {
    if (-not $field.SelectedTypeId -or $field.SelectedTypeId -eq 'Auto') {
        return (TokenTypeLabel ([string]$field.DetectedTypeId))
    }
    if ($field.SelectedTypeId -eq 'DateTime' -and $field.SelectedFormat) {
        return "$(T 'Type_DateTime') ($($field.SelectedFormat))"
    }
    if ([string]$field.SelectedTypeId -like 'Custom:*') {
        return [string]$field.SelectedTypeId
    }
    return (TokenTypeLabel ([string]$field.SelectedTypeId))
}

function GetFieldTypeStatus($field) {
    if ($field.IsAmbiguous -and (-not $field.SelectedTypeId -or $field.SelectedTypeId -eq 'Auto')) {
        return (T 'FieldStatus_Choice')
    }
    if ($field.IsAmbiguous) { return (T 'FieldStatus_Resolved') }
    return (T 'FieldStatus_Detected')
}

function GetFieldCandidateSummary($field) {
    $labels = @($field.CandidateTypes | ForEach-Object {
            if ($_.TypeId -eq 'DateTime' -and $_.Format) {
                "$(T 'Type_DateTime') ($($_.Format))"
            }
            elseif ([string]$_.TypeId -like 'Custom:*' -and $_.DisplayName) {
                [string]$_.DisplayName
            }
            else {
                TokenTypeLabel ([string]$_.TypeId)
            }
        } | Select-Object -Unique)
    if ($labels.Count -eq 0) { return (T 'Type_Text') }
    return ($labels -join ', ')
}

<#
.SYNOPSIS
    Resolves the active TypeId and Format for a field, accounting for manual overrides vs automatic inference.
#>
function GetResolvedFieldType($field) {
    if ($field.SelectedTypeId -and $field.SelectedTypeId -ne 'Auto') {
        return [pscustomobject]@{
            TypeId = [string]$field.SelectedTypeId
            Format = [string]$field.SelectedFormat
        }
    }

    $format = $null
    if ($field.DetectedTypeId -eq 'DateTime') {
        $dateCandidates = @($field.CandidateTypes | Where-Object { $_.TypeId -eq 'DateTime' })
        if ($dateCandidates.Count -eq 1) {
            $format = [string]$dateCandidates[0].Format
        }
    }
    return [pscustomobject]@{
        TypeId = [string]$field.DetectedTypeId
        Format = $format
    }
}

<#
.SYNOPSIS
    Validates field values against resolved type rules before execution.
#>
function ValidateFieldValues([hashtable]$values) {
    foreach ($field in $script:Fields) {
        if ($field.IsVirtual -or -not $values.ContainsKey($field.Name)) { continue }
        $resolvedType = GetResolvedFieldType $field
        if (-not (Test-FNTValueType -Value ([string]$values[$field.Name]) -TypeId $resolvedType.TypeId `
                    -Format $resolvedType.Format -CustomTypeRules @($script:CustomTypeRules))) {
            $formatText = if ($resolvedType.Format) { " ($($resolvedType.Format))" } else { '' }
            throw ((T 'Err_TypeMismatch') -f $field.Name, $values[$field.Name], $resolvedType.TypeId, $formatText)
        }
    }
}

#endregion

#region Theme Management
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class DwmApi {
    [DllImport("dwmapi.dll", PreserveSig = true)]
    public static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);
}
"@ -ErrorAction Ignore

<#
.SYNOPSIS
    Applies Light or Dark WPF palette resource dictionaries and sets DWM window title bar attributes.
#>
function Set-WPFWindowTheme {
    param(
        [Parameter(Mandatory)][System.Windows.Window]$Window,
        [Parameter(Mandatory)][ValidateSet('Light', 'Dark')][string]$Theme
    )

    try {
        $helper = New-Object System.Windows.Interop.WindowInteropHelper($Window)
        if ($helper.Handle -ne [IntPtr]::Zero) {
            $val = if ($Theme -eq 'Dark') { 1 } else { 0 }
            [DwmApi]::DwmSetWindowAttribute($helper.Handle, 20, [ref]$val, 4) | Out-Null
            [DwmApi]::DwmSetWindowAttribute($helper.Handle, 19, [ref]$val, 4) | Out-Null
        }
    }
    catch {}

    $palettes = @{
        'Light' = @{
            'BgWindow'           = '#F5F7FA'
            'BgPanel'            = '#FFFFFF'
            'BgHeader'           = '#E4EAF1'
            'BgControl'          = '#FFFFFF'
            'BgControlHover'     = '#EDF2F7'
            'TextPrimary'        = '#1A202C'
            'TextSecondary'      = '#52606D'
            'TextMuted'          = '#718096'
            'BorderBrush'        = '#CBD5E0'
            'ControlBorder'      = '#CBD5E0'
            'GridAltRow'         = '#EDF2F7'
            'GridRowHover'       = '#E2E8F0'
            'GridHeaderBg'       = '#EDF2F7'
            'AccentColor'        = '#2B6CB0'
            'AccentButtonBg'     = '#D6E9FC'
            'AccentButtonFg'     = '#1A365D'
            'SuccessButtonBg'    = '#C6EFCE'
            'SuccessButtonFg'    = '#006100'
            'RowSuccessBg'       = '#E8F5E9'
            'RowErrorBg'         = '#FFE0E0'
            'RowWarningBg'       = '#FFF3E0'
            'PreviewBoxBg'       = '#FFFEF5'
            'PreviewBoxBorder'   = '#E0D8B0'
            'ListBoxSelectionBg' = '#3182CE'
            'ListBoxSelectionFg' = '#FFFFFF'
            'TabSelectedBg'      = '#FFFFFF'
            'TabSelectedFg'      = '#2B6CB0'
            'TabUnselectedFg'    = '#4A5568'
        }
        'Dark'  = @{
            'BgWindow'           = '#121417'
            'BgPanel'            = '#1E2228'
            'BgHeader'           = '#181B20'
            'BgControl'          = '#282C34'
            'BgControlHover'     = '#323842'
            'TextPrimary'        = '#E6EDF3'
            'TextSecondary'      = '#A0AEC0'
            'TextMuted'          = '#718096'
            'BorderBrush'        = '#3A414D'
            'ControlBorder'      = '#4A5568'
            'GridAltRow'         = '#23272F'
            'GridRowHover'       = '#2D333B'
            'GridHeaderBg'       = '#252930'
            'AccentColor'        = '#63B3ED'
            'AccentButtonBg'     = '#2B4C7E'
            'AccentButtonFg'     = '#EBF8FF'
            'SuccessButtonBg'    = '#1E4620'
            'SuccessButtonFg'    = '#C6F6D5'
            'RowSuccessBg'       = '#17341A'
            'RowErrorBg'         = '#4A1D1D'
            'RowWarningBg'       = '#4A3515'
            'PreviewBoxBg'       = '#1A202C'
            'PreviewBoxBorder'   = '#4A5568'
            'ListBoxSelectionBg' = '#2B6CB0'
            'ListBoxSelectionFg' = '#FFFFFF'
            'TabSelectedBg'      = '#1E2228'
            'TabSelectedFg'      = '#63B3ED'
            'TabUnselectedFg'    = '#A0AEC0'
        }
    }

    $brushConverter = New-Object System.Windows.Media.BrushConverter
    $selectedPalette = $palettes[$Theme]
    foreach ($key in $selectedPalette.Keys) {
        $hex = $selectedPalette[$key]
        $brush = [System.Windows.Media.Brush]($brushConverter.ConvertFromString($hex))
        if ($null -ne $brush) {
            if ($brush.CanFreeze) { $brush.Freeze() }
            $Window.Resources[$key] = $brush
        }
    }
}
#endregion
