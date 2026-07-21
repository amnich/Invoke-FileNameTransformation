# Preview.ps1 — Output preview builder, grid refreshing, and execution logic.
# Dot-sourced by the main script; operates in $script: scope.

function RefreshFieldSelector {
    $names = @($script:Fields | ForEach-Object { $_.Name })
    $prev = $FieldSelector.SelectedItem
    $FieldSelector.ItemsSource = $names
    if ($prev -and $names -contains $prev) {
        $FieldSelector.SelectedItem = $prev
    }
}

function UpdateOutputExample {
    if (-not $script:OutputParts.Count) {
        $OutputExample.Text = (T 'Hint_AddElements')
        return
    }

    # Structure display
    $structParts = @($script:OutputParts | ForEach-Object { $_.Display })
    $structure = ($structParts -join '  +  ')

    # Try live preview from first file
    $preview = ''
    if ($script:CurrentPattern -and $script:CurrentPattern.Items.Count -gt 0) {
        try {
            $item = $script:CurrentPattern.Items[0]
            $values = @{}

            # 1. Extract raw values from filename fields
            foreach ($f in $script:Fields) {
                if (-not $f.IsVirtual -and $f.PartIndex -ge 0) {
                    $values[$f.Name] = $item.Parts[$f.PartIndex].Value
                }
            }
            $meta = Get-FNTFileMetadata -Path $item.File.FullName
            $values[(T 'Name_MetaDate')] = if ($meta.CreationDateStr) { $meta.CreationDateStr } else { '' }
            $values[(T 'Name_MetaAuthor')] = if ($meta.AuthorSegment) { $meta.AuthorSegment } elseif ($meta.Author) { $meta.Author } else { '' }
            ValidateFieldValues $values

            # 2. Apply mappings (in order, allows chaining)
            foreach ($m in $script:Mappings) {
                if (-not (Test-Path $m.Path)) { continue }
                $csv = Import-Csv -LiteralPath $m.Path -Delimiter $m.Delimiter
                $inputVal = [string]$values[$m.InputField]
                $match = $csv | Where-Object {
                    ([string]$_.($m.KeyColumn)).Trim() -eq $inputVal
                } | Select-Object -First 1
                if ($match) {
                    $values[$m.OutputField] = ([string]$match.($m.ValueColumn)).Trim()
                }
            }

            # 3. Apply per-field transforms
            foreach ($f in $script:Fields) {
                if ($values.ContainsKey($f.Name) -and $f.Transforms -and $f.Transforms.Count -gt 0) {
                    $values[$f.Name] = ApplyTransforms $values[$f.Name] $f.Transforms
                }
            }

            # 4. Build output name
            $name = ''
            foreach ($p in $script:OutputParts) {
                if ($p.Type -eq 'Text') {
                    $name += $p.Value
                }
                elseif ($values.ContainsKey($p.Value)) {
                    $name += $values[$p.Value]
                }
                else {
                    $name += "?$($p.Value)?"
                }
            }

            $ext = if ($KeepExtension.IsChecked) { $item.File.Extension }
            else { $NewExtension.Text.Trim() }
            if ($ext -and -not $ext.StartsWith('.')) { $ext = '.' + $ext }

            $preview = "`n`n$(T 'Prefix_Source'):   $($item.File.Name)`n$(T 'Prefix_Result'):   $name$ext"
        }
        catch {
            $preview = "`n`n($(T 'Preview_Unavailable'): $(GetLocalizedErrorMessage $_.Exception))"
        }
    }

    $OutputExample.Text = "$structure$preview"
}

function FullBuildPreview {
    $script:PreviewRows.Clear()

    if (-not $script:CurrentPattern) {
        throw (T 'Err_SelectPattern')
    }
    if (-not $script:OutputParts.Count) {
        throw (T 'Err_AddDest')
    }
    foreach ($field in $script:Fields) {
        if (-not $field.IsVirtual -and $field.IsAmbiguous -and $field.SelectedTypeId -eq 'Auto') {
            throw "$(T 'Err_ResolveAmbiguous') $($field.Name)"
        }
    }

    $src = $SourcePath.Text.Trim()
    $dst = $DestinationPath.Text.Trim()
    if (-not (Test-Path $dst -PathType Container)) {
        throw (T 'Err_DestNotExist')
    }

    # Pre-load all mapping files into hashtables
    $maps = @{}
    foreach ($def in $script:Mappings) {
        if (-not (Test-Path $def.Path)) {
            throw "$(T 'Err_MissMapFile') $($def.Path)"
        }
        $h = @{}
        Import-Csv -LiteralPath $def.Path -Delimiter $def.Delimiter | ForEach-Object {
            $k = ([string]$_.($def.KeyColumn)).Trim()
            if ($h.ContainsKey($k)) {
                throw "$(T 'Err_DupKey') '$k' $(T 'Err_InMap') '$($def.Name)'."
            }
            $h[$k] = ([string]$_.($def.ValueColumn)).Trim()
        }
        $maps[$def.Name] = $h
    }

    # Process each file
    $filesToProcess = if ($EnforcePattern.IsChecked) {
        $all = @()
        foreach ($p in $script:Patterns) { $all += $p.Items }
        $all
    }
    else {
        $script:CurrentPattern.Items
    }

    foreach ($item in $filesToProcess) {
        $row = [ordered]@{
            SourcePath          = $item.File.FullName
            SourceRelative      = $item.File.FullName.Substring($src.TrimEnd('\').Length).TrimStart('\')
            DestinationPath     = ''
            DestinationRelative = ''
            StatusCode          = 'Ready'
            Status              = (T 'Cbo_Ready')
            Details             = ''
        }

        try {
            $values = @{}

            # 1. Extract raw values
            if ($EnforcePattern.IsChecked) {
                $parsedDict = ParseNameByTemplate $item.File.BaseName $script:CurrentPattern.Items[0].Parts
                foreach ($f in $script:Fields) {
                    if (-not $f.IsVirtual -and $f.PartIndex -ge 0) {
                        $values[$f.Name] = $parsedDict[$f.PartIndex]
                    }
                }
            }
            else {
                foreach ($f in $script:Fields) {
                    if (-not $f.IsVirtual -and $f.PartIndex -ge 0) {
                        $values[$f.Name] = $item.Parts[$f.PartIndex].Value
                    }
                }
            }
            $meta = Get-FNTFileMetadata -Path $item.File.FullName
            $values[(T 'Name_MetaDate')] = if ($meta.CreationDateStr) { $meta.CreationDateStr } else { '' }
            $values[(T 'Name_MetaAuthor')] = if ($meta.AuthorSegment) { $meta.AuthorSegment } elseif ($meta.Author) { $meta.Author } else { '' }
            ValidateFieldValues $values

            # 2. Apply mappings (in order)
            foreach ($m in $script:Mappings) {
                $key = [string]$values[$m.InputField]
                if (-not $maps[$m.Name].ContainsKey($key)) {
                    throw "$(T 'Err_MissMap') '$key' $(T 'Err_InMap') '$($m.Name)'."
                }
                $values[$m.OutputField] = $maps[$m.Name][$key]
            }

            # 3. Apply per-field transforms
            foreach ($f in $script:Fields) {
                if ($values.ContainsKey($f.Name) -and $f.Transforms -and $f.Transforms.Count -gt 0) {
                    $values[$f.Name] = ApplyTransforms $values[$f.Name] $f.Transforms
                }
            }

            # 4. Build output name
            $name = ''
            foreach ($p in $script:OutputParts) {
                if ($p.Type -eq 'Text') {
                    $name += $p.Value
                }
                else {
                    if (-not $values.ContainsKey($p.Value)) {
                        throw "$(T 'Err_MissField') '$($p.Value)'."
                    }
                    $name += $values[$p.Value]
                }
            }

            # Validate filename characters
            if ($name.IndexOfAny([IO.Path]::GetInvalidFileNameChars()) -ge 0) {
                throw ($(T 'Err_InvalidChars') + " $name")
            }

            # Extension
            $ext = if ($KeepExtension.IsChecked) { $item.File.Extension }
            else { $NewExtension.Text.Trim() }
            if ($ext -and -not $ext.StartsWith('.')) { $ext = '.' + $ext }

            # Build destination path
            $relativeDir = if ($PreserveFolderStructure.IsChecked) { Split-Path $row.SourceRelative -Parent } else { '' }
            $destDir = if ($relativeDir) { Join-Path $dst $relativeDir } else { $dst }
            $row.DestinationPath = Join-Path $destDir ($name + $ext)
            $row.DestinationRelative = $row.DestinationPath.Substring($dst.TrimEnd('\').Length).TrimStart('\')

            # Details: show all field values
            $row.Details = ($values.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join '; '

        }
        catch {
            $row.StatusCode = 'Error'
            $row.Status = (T 'Title_Error')
            $row.Details = GetLocalizedErrorMessage $_.Exception
        }

        $script:PreviewRows.Add([pscustomobject]$row)
    }

    # Check for duplicate destinations
    $dupes = @(
        $script:PreviewRows |
        Where-Object { $_.DestinationPath } |
        Group-Object DestinationPath |
        Where-Object { $_.Count -gt 1 } |
        Select-Object -ExpandProperty Name
    )
    foreach ($r in $script:PreviewRows) {
        if ($r.DestinationPath -in $dupes) {
            $r.StatusCode = 'Error'
            $r.Status = (T 'Title_Error')
            $r.Details = (T 'Err_DupDest')
        }
        elseif ($r.DestinationPath -and (Test-Path $r.DestinationPath)) {
            $r.StatusCode = 'Error'
            $r.Status = (T 'Title_Error')
            $r.Details = (T 'Err_FileExists')
        }
    }

    RefreshPreviewGrid
}

function RefreshPreviewGrid {
    $filter = [string]$PreviewFilter.SelectedItem.Content
    $data = switch ($filter) {
        (T 'Cbo_Errors') { @($script:PreviewRows | Where-Object { $_.StatusCode -eq 'Error' }) }
        (T 'Cbo_Ready') { @($script:PreviewRows | Where-Object { $_.StatusCode -eq 'Ready' }) }
        default { @($script:PreviewRows) }
    }
    $PreviewGrid.ItemsSource = $null
    $PreviewGrid.ItemsSource = $data

    $errCount = @($script:PreviewRows | Where-Object { $_.StatusCode -eq 'Error' }).Count
    $PreviewInfo.Text = "$(T 'Msg_Files'): $($script:PreviewRows.Count); $(T 'Msg_Errors'): $errCount"
}

function ExecuteCopy {
    # Rebuild preview to get fresh state
    FullBuildPreview

    $errors = @($script:PreviewRows | Where-Object { $_.StatusCode -eq 'Error' })
    if ($errors) {
        throw "$(T 'Err_Blocked') $($errors.Count) $(T 'Err_FixPrev')"
    }

    $mode = if ($ExecutionMode -and $ExecutionMode.SelectedIndex -eq 1) { 'Move' } else { 'Copy' }
    $modeLabel = if ($mode -eq 'Move') { (T 'Action_Move') } else { (T 'Action_Copy') }
    $confirmText = if ($mode -eq 'Move') {
        "$modeLabel $((($script:PreviewRows | Where-Object { $_.StatusCode -eq 'Ready' }).Count)) $(T 'Msg_MoveOrig')"
    }
    else {
        "$modeLabel $((($script:PreviewRows | Where-Object { $_.StatusCode -eq 'Ready' }).Count)) $(T 'Msg_CopyOrig')"
    }

    $total = ($script:PreviewRows | Where-Object { $_.StatusCode -eq 'Ready' }).Count
    $confirm = [Windows.MessageBox]::Show(
        $confirmText,
        (T 'Title_Confirm'), 'YesNo', 'Question'
    )
    if ($confirm -ne 'Yes') { return }

    # Show progress bar
    $ProgressBar.Visibility = 'Visible'
    $ProgressBar.Maximum = $total
    $ProgressBar.Value = 0

    $ok = 0; $fail = 0

    foreach ($r in $script:PreviewRows | Where-Object { $_.StatusCode -eq 'Ready' }) {
        try {
            $dir = Split-Path $r.DestinationPath -Parent
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
            if ($mode -eq 'Move') {
                Move-Item -LiteralPath $r.SourcePath -Destination $r.DestinationPath -ErrorAction Stop -Force
                Log "$(T 'Log_Moved') $($r.SourcePath) -> $($r.DestinationPath)"
            }
            else {
                Copy-Item -LiteralPath $r.SourcePath -Destination $r.DestinationPath -ErrorAction Stop -Force
                Log "$(T 'Log_Copied') $($r.SourcePath) -> $($r.DestinationPath)"
            }
            $ok++
        }
        catch {
            $fail++
            Log "$(T 'Log_CopyErr') $($_.Exception.Message)" 'ERROR'
        }
        $ProgressBar.Value = $ok + $fail
        UpdateUI
    }

    $ProgressBar.Visibility = 'Collapsed'
    if ($mode -eq 'Move') {
        SetStatus "$(T 'Status_Moved') $ok; $(T 'Msg_Errors'): $fail"
    }
    else {
        SetStatus "$(T 'Status_Copied') $ok; $(T 'Msg_Errors'): $fail"
    }
}
