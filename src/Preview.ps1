# Preview.ps1 — Output preview builder, grid refreshing, and execution logic.
# Dot-sourced by the main script; operates in $script: scope.

<#
.SYNOPSIS
    Reads data records from CSV, TXT, JSON, or XML lookup files into standard object collections.

.PARAMETER path
    Path to the dictionary data file.

.PARAMETER delimiter
    Delimiter character if reading CSV/TXT files.

.OUTPUTS
    [PSCustomObject[]] Array of data record objects.
#>
function Read-FNTDictionaryData([string]$path, [string]$delimiter) {
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { return @() }
    $ext = [System.IO.Path]::GetExtension($path).ToLowerInvariant()

    if ($ext -eq '.json') {
        try {
            $raw = Get-Content -LiteralPath $path -Raw -Encoding UTF8
            $json = $raw | ConvertFrom-Json
            if ($json -is [System.Array]) { return @($json) }
            elseif ($json -is [PSCustomObject]) { return @($json) }
        } catch { return @() }
    }
    elseif ($ext -eq '.xml') {
        try {
            [xml]$xml = Get-Content -LiteralPath $path -Raw -Encoding UTF8
            $root = $xml.DocumentElement
            $list = New-Object System.Collections.Generic.List[pscustomobject]
            foreach ($child in $root.ChildNodes) {
                $obj = [ordered]@{}
                foreach ($sub in $child.ChildNodes) {
                    $obj[$sub.Name] = $sub.InnerText
                }
                foreach ($attr in $child.Attributes) {
                    $obj["@" + $attr.Name] = $attr.Value
                }
                $list.Add([pscustomobject]$obj)
            }
            return $list.ToArray()
        } catch { return @() }
    }

    return @(Import-Csv -LiteralPath $path -Delimiter $delimiter)
}

<#
.SYNOPSIS
    Refreshes the dropdown list of available field names in Tab 4.
#>
function RefreshFieldSelector {
    $names = @($script:Fields | ForEach-Object { $_.Name })
    $prev = $FieldSelector.SelectedItem
    $FieldSelector.ItemsSource = $names
    if ($prev -and $names -contains $prev) {
        $FieldSelector.SelectedItem = $prev
    }
}

<#
.SYNOPSIS
    Updates the live sample output string box in Tab 4 using the first file in the active pattern group.
#>
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
    $patternItems = @($script:CurrentPattern.Items)
    if ($script:CurrentPattern -and $patternItems.Count -gt 0) {
        try {
            $item = $patternItems[0]
            $values = @{}

            # 1. Extract raw values from filename fields
            foreach ($f in $script:Fields) {
                if (-not $f.IsVirtual -and $f.PartIndex -ge 0) {
                    $values[$f.Name] = $item.Parts[$f.PartIndex].Value
                }
            }
            $reqFields = @($script:OutputParts | ForEach-Object { $_.Value })
            $meta = Get-FNTFileMetadata -Path $item.File.FullName -RequestedFields $reqFields
            $values[(T 'Name_MetaDate')]       = if ($meta.CreationDateStr) { $meta.CreationDateStr } else { '' }
            $values[(T 'Name_MetaAuthor')]     = if ($meta.AuthorSegment) { $meta.AuthorSegment } elseif ($meta.Author) { $meta.Author } else { '' }
            $values[(T 'Name_MetaTitle')]      = if ($meta.Title) { $meta.Title } else { '' }
            $values[(T 'Name_MetaDateTaken')]  = if ($meta.DateTakenStr) { $meta.DateTakenStr } else { '' }
            $values[(T 'Name_MetaDimensions')] = if ($meta.Dimensions) { $meta.Dimensions } else { '' }
            $values[(T 'Name_MetaCamera')]     = if ($meta.Camera) { $meta.Camera } else { '' }
            $values[(T 'Name_MetaAudioArtist')]= if ($meta.AudioArtist) { $meta.AudioArtist } else { '' }
            $values[(T 'Name_MetaDocCreator')] = if ($meta.DocCreator) { $meta.DocCreator } else { '' }
            $values[(T 'Name_MetaHashMD5')]    = if ($meta.HashMD5) { $meta.HashMD5 } else { '' }
            $values[(T 'Name_MetaHashSHA256')] = if ($meta.HashSHA256) { $meta.HashSHA256 } else { '' }
            ValidateFieldValues $values

            # 2. Apply mappings (in order, allows chaining)
            foreach ($m in $script:Mappings) {
                if (-not (Test-Path $m.Path)) { continue }
                $data = Read-FNTDictionaryData -path $m.Path -delimiter $m.Delimiter
                $inputVal = [string]$values[$m.InputField]
                $match = $data | Where-Object {
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
        $data = Read-FNTDictionaryData -path $def.Path -delimiter $def.Delimiter
        $data | ForEach-Object {
            $k = ([string]$_.($def.KeyColumn)).Trim()
            if ($h.ContainsKey($k)) {
                throw "$(T 'Err_DupKey') '$k' $(T 'Err_InMap') '$($def.Name)'."
            }
            $h[$k] = ([string]$_.($def.ValueColumn)).Trim()
        }
        $maps[$def.Name] = $h
    }

    $reqFields = @($script:OutputParts | ForEach-Object { $_.Value })
    $collisionPolicy = 'Block'
    if ($CollisionPolicySelector -and $CollisionPolicySelector.SelectedItem) {
        $tagVal = [string]$CollisionPolicySelector.SelectedItem.Tag
        if ($tagVal) { $collisionPolicy = $tagVal }
    }
    $claimedPaths = @{}

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
            $meta = Get-FNTFileMetadata -Path $item.File.FullName -RequestedFields $reqFields
            $values[(T 'Name_MetaDate')]       = if ($meta.CreationDateStr) { $meta.CreationDateStr } else { '' }
            $values[(T 'Name_MetaAuthor')]     = if ($meta.AuthorSegment) { $meta.AuthorSegment } elseif ($meta.Author) { $meta.Author } else { '' }
            $values[(T 'Name_MetaTitle')]      = if ($meta.Title) { $meta.Title } else { '' }
            $values[(T 'Name_MetaDateTaken')]  = if ($meta.DateTakenStr) { $meta.DateTakenStr } else { '' }
            $values[(T 'Name_MetaDimensions')] = if ($meta.Dimensions) { $meta.Dimensions } else { '' }
            $values[(T 'Name_MetaCamera')]     = if ($meta.Camera) { $meta.Camera } else { '' }
            $values[(T 'Name_MetaAudioArtist')]= if ($meta.AudioArtist) { $meta.AudioArtist } else { '' }
            $values[(T 'Name_MetaDocCreator')] = if ($meta.DocCreator) { $meta.DocCreator } else { '' }
            $values[(T 'Name_MetaHashMD5')]    = if ($meta.HashMD5) { $meta.HashMD5 } else { '' }
            $values[(T 'Name_MetaHashSHA256')] = if ($meta.HashSHA256) { $meta.HashSHA256 } else { '' }
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

            # Extension
            $ext = if ($KeepExtension.IsChecked) { $item.File.Extension }
            else { $NewExtension.Text.Trim() }
            if ($ext -and -not $ext.StartsWith('.')) { $ext = '.' + $ext }

            # Build destination path & check collisions
            $relativeDir = if ($PreserveFolderStructure.IsChecked) { Split-Path $row.SourceRelative -Parent } else { '' }
            $destDir = if ($relativeDir) { Join-Path $dst $relativeDir } else { $dst }
            $rawDestPath = Join-Path $destDir ($name + $ext)

            $res = Resolve-FNTDestinationCollision -DestinationPath $rawDestPath -CollisionPolicy $collisionPolicy -ClaimedPaths $claimedPaths
            if ($res.Action -eq 'Skip') {
                $row.StatusCode = 'Skipped'
                $row.Status = (T 'Policy_Skip')
                $row.Details = (T 'Policy_Skip')
            }
            else {
                $row.DestinationPath = $res.Path
                $row.DestinationRelative = $row.DestinationPath.Substring($dst.TrimEnd('\').Length).TrimStart('\')
                $detailsList = ($values.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join '; '
                if ($res.Action -eq 'AutoNumber') {
                    $detailsList = "[AutoNumber] $detailsList"
                }
                elseif ($res.Action -eq 'Overwrite') {
                    $detailsList = "[Overwrite] $detailsList"
                }
                $row.Details = $detailsList
            }

        }
        catch {
            $row.StatusCode = 'Error'
            $row.Status = (T 'Title_Error')
            $row.Details = GetLocalizedErrorMessage $_.Exception
        }

        $script:PreviewRows.Add([pscustomobject]$row)
    }

    # Check for duplicate destinations if Block policy
    if ($collisionPolicy -eq 'Block') {
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
    $completedOps = New-Object System.Collections.Generic.List[object]

    foreach ($r in $script:PreviewRows | Where-Object { $_.StatusCode -eq 'Ready' }) {
        try {
            $dir = Split-Path $r.DestinationPath -Parent
            if (-not (Test-Path -LiteralPath $dir -PathType Container)) {
                New-Item -ItemType Directory -Path $dir -Force | Out-Null
            }
            if ($mode -eq 'Move') {
                Move-Item -LiteralPath $r.SourcePath -Destination $r.DestinationPath -ErrorAction Stop -Force
                Log "$(T 'Log_Moved') $($r.SourcePath) -> $($r.DestinationPath)"
            }
            else {
                Copy-Item -LiteralPath $r.SourcePath -Destination $r.DestinationPath -ErrorAction Stop -Force
                Log "$(T 'Log_Copied') $($r.SourcePath) -> $($r.DestinationPath)"
            }
            $completedOps.Add([pscustomobject]@{
                Mode            = $mode
                SourcePath      = $r.SourcePath
                DestinationPath = $r.DestinationPath
            })
            $ok++
        }
        catch {
            $fail++
            Log "$(T 'Log_CopyErr') $($_.Exception.Message)" 'ERROR'
        }
        $ProgressBar.Value = $ok + $fail
        UpdateUI
    }

    if ($completedOps.Count -gt 0) {
        try {
            $manifestPath = Export-FNTUndoManifest -Operations $completedOps.ToArray() -LogDirectory $script:LogRoot
            Log "Undo manifest created: $manifestPath"
        }
        catch {
            Log "Failed to create undo manifest: $($_.Exception.Message)" 'WARNING'
        }
    }

    $ProgressBar.Visibility = 'Collapsed'
    if ($mode -eq 'Move') {
        SetStatus "$(T 'Status_Moved') $ok; $(T 'Msg_Errors'): $fail"
    }
    else {
        SetStatus "$(T 'Status_Copied') $ok; $(T 'Msg_Errors'): $fail"
    }
}

function Invoke-FNTUndoLastOperation {
    if (-not $script:LogRoot -or -not (Test-Path -LiteralPath $script:LogRoot -PathType Container)) {
        throw (T 'Title_MissingData')
    }
    $latestManifest = Get-ChildItem -LiteralPath $script:LogRoot -Filter 'undo_*.json' | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if (-not $latestManifest) {
        throw (T 'Title_MissingData')
    }
    $confirm = [Windows.MessageBox]::Show(
        "Revert last operation from '$($latestManifest.Name)'?",
        (T 'Title_Confirm'), 'YesNo', 'Question'
    )
    if ($confirm -ne 'Yes') { return }

    $res = Invoke-FNTUndoOperation -ManifestPath $latestManifest.FullName
    if ($res.FailedCount -eq 0) {
        SetStatus "$(T 'Status_UndoDone') $($res.RevertedCount)"
    }
    else {
        SetStatus "$(T 'Status_UndoErr') Reverted: $($res.RevertedCount), Failed: $($res.FailedCount)"
    }
    Remove-Item -LiteralPath $latestManifest.FullName -Force -ErrorAction SilentlyContinue
}
