# UI.ps1 — Utility functions, logging, dialogs, and field-type helpers.
# Dot-sourced by the main script; operates in $script: scope.

#region Utility Functions

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

function ValidateFieldValues([hashtable]$values) {
    foreach ($field in $script:Fields) {
        if ($field.IsVirtual -or -not $values.ContainsKey($field.Name)) { continue }
        $resolvedType = GetResolvedFieldType $field
        if (-not (Test-FNTValueType -Value ([string]$values[$field.Name]) -TypeId $resolvedType.TypeId `
                    -Format $resolvedType.Format -CustomTypeRules @($script:CustomTypeRules))) {
            $formatText = if ($resolvedType.Format) { " ($($resolvedType.Format))" } else { '' }
            throw "$($field.Name): '$($values[$field.Name])' != $($resolvedType.TypeId)$formatText"
        }
    }
}

#endregion
