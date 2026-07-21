# Compliance.ps1 — Naming convention compliance scanner, fixer, and metadata integration.
# Dot-sourced by the main script; operates in $script: scope.

$script:ComplianceRows = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'

function BuildComplianceSuggestion($item, $complianceResult, $metadata) {
    # Base parts: Date_Author_DocType_FreeText_Version
    $seg = $complianceResult.Segments
    $authorInfo = $complianceResult.AuthorAnalysis

    # 1. Date: use extracted metadata creation date, or parsed date, or file creation date
    $datePart = if ($seg.Date -match '^\d{8}$') { $seg.Date }
    elseif ($metadata.CreationDateStr) { $metadata.CreationDateStr }
    else { (Get-Date).ToString('yyyyMMdd') }

    # 2. Author: use suggested author or normalized parsed author if non-empty, else metadata AuthorSegment, else default
    $authorPart = if ($authorInfo.SuggestedAuthor -and $authorInfo.SuggestedAuthor.Length -eq 8) { $authorInfo.SuggestedAuthor }
    elseif ($seg.Author) { Get-FNTNormalizedAuthorSegment $seg.Author }
    elseif ($metadata.AuthorSegment -and $metadata.AuthorSegment.Length -eq 8) { $metadata.AuthorSegment }
    else { 'UnknownA' }

    # 3. DocType: use parsed DocType or default keyword
    $docTypePart = if ($seg.DocType) { $seg.DocType } else { 'Dokument' }

    # 4. FreeText: use parsed FreeText or BaseName sanitized
    $freeTextPart = if ($seg.FreeText) { $seg.FreeText }
    else {
        # Fallback: clean original basename
        $clean = $item.BaseName -replace '[_\-\s]+', '-'
        if ($clean.Length -gt 30) { $clean = $clean.Substring(0, 30) }
        $clean
    }

    # 5. Version: preserve if present
    $versionPart = if ($seg.Version) { "_$($seg.Version)" } else { '' }

    return "${datePart}_${authorPart}_${docTypePart}_${freeTextPart}${versionPart}$($item.Extension)"
}

function ScanCompliance {
    $src = $SourcePath.Text.Trim()
    if (-not (Test-Path $src -PathType Container)) {
        throw (T 'Err_SrcNotExist')
    }

    $files = if ($ComplianceRecursive.IsChecked) {
        @(Get-ChildItem -LiteralPath $src -File -Recurse)
    }
    else {
        @(Get-ChildItem -LiteralPath $src -File)
    }

    if (-not $files) { throw (T 'Err_NoFiles') }

    # Extension filter population
    $extensions = @($files | ForEach-Object { $_.Extension.ToLower() } | Sort-Object -Unique)
    $ComplianceExtFilter.ItemsSource = $extensions
    $selectedExt = [string]$ComplianceExtFilter.SelectedItem

    $filteredFiles = if ($selectedExt) {
        @($files | Where-Object { $_.Extension.ToLower() -eq $selectedExt })
    } else {
        $files
    }

    $script:ComplianceRows.Clear()
    $okCount = 0
    $failCount = 0

    foreach ($file in $filteredFiles) {
        $meta = Get-FNTFileMetadata -Path $file.FullName
        $res = Test-FNTNamingConvention -BaseName $file.BaseName

        $violationsText = if ($res.Violations) {
            (@($res.Violations | ForEach-Object { T $_ }) -join '; ')
        } else { '' }

        $suggestedName = if (-not $res.IsCompliant) {
            BuildComplianceSuggestion $file $res $meta
        } else { $file.Name }

        if ($res.IsCompliant) { $okCount++ } else { $failCount++ }

        $row = [pscustomobject]@{
            File           = $file
            FileName       = $file.Name
            FullName       = $file.FullName
            IsCompliant    = $res.IsCompliant
            StatusCode     = if ($res.IsCompliant) { 'Ready' } else { 'Error' }
            StatusText     = if ($res.IsCompliant) { T 'Compliance_OK' } else { T 'Compliance_Fail' }
            MetaDate       = if ($meta.CreationDateStr) { $meta.CreationDateStr } else { T 'Compliance_MetaUnavailable' }
            MetaAuthor     = if ($meta.Author) { $meta.Author } else { T 'Compliance_MetaUnavailable' }
            AuthorSegment  = $meta.AuthorSegment
            ViolationText  = $violationsText
            SuggestedName  = $suggestedName
            Details        = "Path: $($file.FullName); Violations: $violationsText"
        }

        $script:ComplianceRows.Add($row)
    }

    $ComplianceGrid.ItemsSource = $null
    $ComplianceGrid.ItemsSource = $script:ComplianceRows
    $ComplianceInfo.Text = "$(T 'Msg_Files'): $($filteredFiles.Count); $(T 'Msg_Compliant'): $okCount; $(T 'Msg_NonCompliant'): $failCount"
}

function ApplyComplianceFix {
    $selected = @($ComplianceGrid.SelectedItems)

    # If no rows selected, select all rows with suggested name changes
    if (-not $selected -or $selected.Count -eq 0) {
        $selected = @($script:ComplianceRows | Where-Object { $_.SuggestedName -and $_.SuggestedName -ne $_.FileName })
    }

    # Filter to items that actually need renaming
    $itemsToRename = @($selected | Where-Object { $_.SuggestedName -and $_.SuggestedName -ne $_.FileName })
    if (-not $itemsToRename -or $itemsToRename.Count -eq 0) { return }

    $count = $itemsToRename.Count
    $msgTemplate = T 'Msg_ConfirmRenameInPlace'
    $confirmMsg = if ($msgTemplate -and $msgTemplate -ne 'Msg_ConfirmRenameInPlace') {
        $msgTemplate -f $count
    } else {
        "Zmienić nazwy $count plików w miejscu na sugerowane nazwy?"
    }

    $confirm = [Windows.MessageBox]::Show(
        $confirmMsg,
        (T 'Title_Confirm'), 'YesNo', 'Question'
    )
    if ($confirm -ne 'Yes') { return }

    $fixed = 0
    foreach ($item in $itemsToRename) {
        try {
            $dir = Split-Path $item.FullName -Parent
            $newPath = Join-Path $dir $item.SuggestedName
            if (-not (Test-Path -LiteralPath $newPath)) {
                Rename-Item -LiteralPath $item.FullName -NewName $item.SuggestedName -ErrorAction Stop
                Log "Renamed: $($item.FullName) -> $($item.SuggestedName)"
                $fixed++
            }
        }
        catch {
            Log "Compliance Rename Error ($($item.FileName)): $($_.Exception.Message)" 'ERROR'
        }
    }

    $statusTemplate = T 'Status_RenamedCount'
    $statusMsg = if ($statusTemplate -and $statusTemplate -ne 'Status_RenamedCount') {
        $statusTemplate -f $fixed
    } else {
        "Zmieniono nazwy $fixed plików w miejscu."
    }

    SetStatus $statusMsg
    ScanCompliance
}

function InjectMetadataVirtualFields {
    EnsureVirtualField (T 'Name_MetaDate')
    EnsureVirtualField (T 'Name_MetaAuthor')
}
