# Analysis.ps1 — Tokenization, pattern analysis, and field detection.
# Dot-sourced by the main script; operates in $script: scope.

function Tokens([string]$name) {
    $pattern = if ($TokenRegex -and $TokenRegex.Text) { $TokenRegex.Text } else { '(?<value>[^_\-\s]+)|(?<sep>[_\-\s]+)' }
    Get-FNTTokens -Name $name -Pattern $pattern -CustomTypeRules @($script:CustomTypeRules)
}

function ParseNameByTemplate([string]$name, $templateParts) {
    $fieldTypes = @{}
    foreach ($field in $script:Fields) {
        if (-not $field.IsVirtual -and $field.PartIndex -ge 0) {
            $fieldTypes[[int]$field.PartIndex] = GetResolvedFieldType $field
        }
    }
    $pattern = if ($TokenRegex -and $TokenRegex.Text) { $TokenRegex.Text } else { '(?<value>[^_\-\s]+)|(?<sep>[_\-\s]+)' }
    $match = Match-FNTNamePattern -Name $name -PatternTokens @($templateParts) -TokenizerPattern $pattern `
        -FieldTypes $fieldTypes -CustomTypeRules @($script:CustomTypeRules)
    return $match.Values
}

function AnalyzePatterns {
    $src = $SourcePath.Text.Trim()
    if (-not (Test-Path $src -PathType Container)) {
        throw (T 'Err_SrcNotExist')
    }

    $files = if ($Recursive.IsChecked) {
        @(Get-ChildItem -LiteralPath $src -File -Recurse)
    }
    else {
        @(Get-ChildItem -LiteralPath $src -File)
    }
    if (-not $files) { throw (T 'Err_NoFiles') }

    # Populate extension filter
    $extensions = @($files | ForEach-Object { $_.Extension.ToLower() } | Sort-Object -Unique)
    $ExtensionFilter.ItemsSource = $extensions
    if (-not $ExtensionFilter.SelectedItem) {
        $ExtensionFilter.SelectedItem = ($extensions | Select-Object -First 1)
    }

    BuildPatternList
}

function BuildPatternList {
    $src = $SourcePath.Text.Trim()
    $ext = [string]$ExtensionFilter.SelectedItem
    if (-not $ext) { return }

    $files = if ($Recursive.IsChecked) {
        @(Get-ChildItem -LiteralPath $src -File -Recurse -Filter "*$ext")
    }
    else {
        @(Get-ChildItem -LiteralPath $src -File -Filter "*$ext")
    }

    # Group files by token structure signature
    $raw = @()
    foreach ($f in $files) {
        $parts = Tokens $f.BaseName
        $sig = Get-FNTTokenSignature -Tokens @($parts)
        $raw += [pscustomobject]@{ Signature = $sig; Parts = $parts; File = $f }
    }

    $script:Patterns = @()
    foreach ($g in $raw | Group-Object Signature) {
        $sample = $g.Group[0]
        $all = $g.Group
        $labels = @()
        $fieldInferences = @{}
        for ($i = 0; $i -lt $sample.Parts.Count; $i++) {
            if ($sample.Parts[$i].IsSeparator) {
                $labels += $sample.Parts[$i].Value
                continue
            }
            $uniqueValues = @($all | ForEach-Object { $_.Parts[$i].Value } | Select-Object -Unique)
            $inference = Get-FNTFieldInference -Tokens @($all | ForEach-Object { $_.Parts[$i] })
            $fieldInferences[$i] = $inference
            $kind = TokenTypeLabel $inference.DetectedTypeId
            $labels += if ($uniqueValues.Count -eq 1) { "[$($uniqueValues[0])]" } else { "<$kind>" }
        }
        $script:Patterns += [pscustomobject]@{
            Extension       = $ext
            Display         = ($labels -join '')
            Count           = $g.Count
            Items           = $all
            Signature       = $g.Name
            FieldInferences = $fieldInferences
        }
    }

    $PatternGrid.ItemsSource = $script:Patterns
    $AnalysisInfo.Text = "$(T 'Msg_Files'): $($files.Count); $(T 'Msg_Structs'): $($script:Patterns.Count)"
}

function SetPattern($pattern) {
    $script:CurrentPattern = $pattern

    # Show sample file names
    $PatternSamples.ItemsSource = @(
        $pattern.Items | Select-Object -First 20 | ForEach-Object { $_.File.Name }
    )

    # Build fields from token positions (clear old fields)
    $script:Fields.Clear()
    $parts = $pattern.Items[0].Parts
    for ($i = 0; $i -lt $parts.Count; $i++) {
        if ($parts[$i].IsSeparator) { continue }

        $uniqueValues = @($pattern.Items | ForEach-Object { $_.Parts[$i].Value } | Select-Object -Unique)
        $inference = $pattern.FieldInferences[$i]
        $typeId = $inference.DetectedTypeId
        $type = TokenTypeLabel $typeId

        # Auto-detect role
        $role = if ($typeId -eq 'DateTime') { (T 'Role_Date') }
        elseif ($uniqueValues.Count -eq 1) { (T 'Role_Const') }
        else { (T 'Role_Value') }

        # Auto-name
        $name = switch ($role) {
            (T 'Role_Date') { "$(T 'Name_Date')_$($i + 1)" }
            (T 'Role_Const') { "$(T 'Name_Text')_$($i + 1)" }
            default { "$(T 'Name_Field')_$($i + 1)" }
        }

        $script:Fields.Add([pscustomobject]@{
                PartIndex        = $i
                DisplayIndex     = "$($script:Fields.Count + 1)"
                Sample           = $parts[$i].Value
                DetectedType     = $type
                DetectedTypeId   = $typeId
                CandidateTypes   = @($inference.CandidateTypes)
                IsAmbiguous      = [bool]$inference.IsAmbiguous
                SelectedTypeId   = 'Auto'
                SelectedFormat   = $null
                EffectiveType    = $type
                TypeStatus       = if ($inference.IsAmbiguous) { (T 'FieldStatus_Choice') } else { (T 'FieldStatus_Detected') }
                CandidateSummary = GetFieldCandidateSummary ([pscustomobject]@{ CandidateTypes = @($inference.CandidateTypes) })
                Name             = $name
                Role             = $role
                IsVirtual        = $false
                Source           = (T 'Src_Name')
                Transforms       = [System.Collections.ArrayList]::new()
            })
    }

    # Recreate virtual fields from existing mappings and file metadata
    foreach ($m in $script:Mappings) {
        EnsureVirtualField $m.OutputField
    }
    EnsureVirtualField (T 'Name_MetaDate')
    EnsureVirtualField (T 'Name_MetaAuthor')

    $FieldGrid.ItemsSource = $script:Fields
    RefreshFieldSelector
    $PatternHint.Text = (T 'Hint_SelectField')
    UpdateOutputExample
}
