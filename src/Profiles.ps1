# Profiles.ps1 — Saving, loading, and refreshing JSON profiles.
# Dot-sourced by the main script; operates in $script: scope.

<#
.SYNOPSIS
    Enumerates saved JSON profiles in the profile root directory and updates the UI profile list.
#>
function RefreshProfiles {
    $items = @(
        Get-ChildItem $script:ProfileRoot -Filter '*.json' -ErrorAction SilentlyContinue |
        ForEach-Object { [pscustomobject]@{ Name = $_.BaseName; Path = $_.FullName } }
    )
    $ProfileList.ItemsSource = $items
}

<#
.SYNOPSIS
    Prompts for a profile name and saves the current field, mapping, and output setup to a JSON profile file (Schema V2).
#>
function SaveProfile {
    $name = [Microsoft.VisualBasic.Interaction]::InputBox(
        (T 'Lbl_ProfileName'), (T 'Title_SaveProfile'), $script:CurrentProfileName
    )
    if (-not $name) { return }

    # Serialize fields with transforms converted to plain arrays
    $fieldsData = @($script:Fields | ForEach-Object {
            $clone = [ordered]@{}
            $_.PSObject.Properties | ForEach-Object { $clone[$_.Name] = $_.Value }
            $clone.Transforms = @($_.Transforms)
            [pscustomobject]$clone
        })

    $obj = [ordered]@{
        SchemaVersion = 2
        Name          = $name
        TokenRegex    = $TokenRegex.Text
        Fields        = $fieldsData
        Mappings      = @($script:Mappings)
        OutputParts   = @($script:OutputParts)
        KeepExtension = [bool]$KeepExtension.IsChecked
        NewExtension  = $NewExtension.Text
        FolderPattern = if ($FolderPattern) { $FolderPattern.Text } else { '' }
    }
    $obj = ConvertTo-FNTProfile ([pscustomobject]$obj)

    $path = Join-Path $script:ProfileRoot ($name + '.json')
    $obj | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $path -Encoding UTF8

    $script:CurrentProfileName = $name
    $CurrentProfile.Text = $name
    RefreshProfiles
    SetStatus "$(T 'Status_ProfSaved') $name"
}

<#
.SYNOPSIS
    Loads and normalizes a JSON profile file, restoring fields, transforms, mappings, and output name components.

.PARAMETER path
    Path to the JSON profile file.
#>
function LoadProfile([string]$path) {
    $rawProfile = Get-Content -LiteralPath $path -Raw -Encoding UTF8 | ConvertFrom-Json
    $p = ConvertTo-FNTProfile $rawProfile

    # Restore fields
    $script:Fields.Clear()
    @($p.Fields) | ForEach-Object {
        # Ensure Transforms is ArrayList
        $transforms = [System.Collections.ArrayList]::new()
        if ($_.PSObject.Properties['Transforms'] -and $_.Transforms) {
            @($_.Transforms) | ForEach-Object { [void]$transforms.Add($_) }
        }
        $_.Transforms = $transforms

        if (-not $_.PSObject.Properties['Preview']) {
            $_ | Add-Member -NotePropertyName 'Preview' -NotePropertyValue '' -Force
        }

        if (-not $_.PSObject.Properties['Source']) {
            $src = if ($_.IsVirtual) { (T 'Src_Mapping') } else { (T 'Src_Name') }
            $_ | Add-Member -NotePropertyName 'Source' -NotePropertyValue $src -Force
        }
        $_.DetectedType = TokenTypeLabel ([string]$_.DetectedTypeId)
        if (-not $_.PSObject.Properties['EffectiveType']) {
            $_ | Add-Member -NotePropertyName 'EffectiveType' -NotePropertyValue (GetEffectiveTypeLabel $_) -Force
        }
        else {
            $_.EffectiveType = GetEffectiveTypeLabel $_
        }
        if (-not $_.PSObject.Properties['TypeStatus']) {
            $_ | Add-Member -NotePropertyName 'TypeStatus' -NotePropertyValue (GetFieldTypeStatus $_) -Force
        }
        else {
            $_.TypeStatus = GetFieldTypeStatus $_
        }
        if (-not $_.PSObject.Properties['CandidateSummary']) {
            $_ | Add-Member -NotePropertyName 'CandidateSummary' -NotePropertyValue (GetFieldCandidateSummary $_) -Force
        }
        else {
            $_.CandidateSummary = GetFieldCandidateSummary $_
        }

        $script:Fields.Add($_)
    }

    # Restore mappings
    $script:Mappings.Clear()
    @($p.Mappings) | ForEach-Object {
        if (-not $_.PSObject.Properties['Display']) {
            $_ | Add-Member -NotePropertyName 'Display' -NotePropertyValue "$($_.Name): $($_.InputField) → $($_.OutputField)" -Force
        }
        $script:Mappings.Add($_)
    }

    # Restore output parts
    $script:OutputParts.Clear()
    @($p.OutputParts) | ForEach-Object { $script:OutputParts.Add($_) }

    # Restore extension settings
    $KeepExtension.IsChecked = $p.KeepExtension
    $NewExtension.Text = if ($p.NewExtension) { $p.NewExtension } else { '' }
    if ($p.PSObject.Properties['TokenRegex'] -and -not [string]::IsNullOrWhiteSpace([string]$p.TokenRegex)) {
        $TokenRegex.Text = [string]$p.TokenRegex
    }
    if ($FolderPattern) {
        $FolderPattern.Text = if ($p.PSObject.Properties['FolderPattern'] -and $p.FolderPattern) { [string]$p.FolderPattern } else { '' }
    }

    # Refresh UI bindings
    $FieldGrid.ItemsSource = $script:Fields
    $MappingList.ItemsSource = $script:Mappings
    $OutputList.ItemsSource = $script:OutputParts
    RefreshFieldSelector
    UpdateOutputExample

    $script:CurrentProfileName = $p.Name
    $CurrentProfile.Text = $p.Name
    SetStatus "$(T 'Status_ProfLoaded') $($p.Name)"
}

<#
.SYNOPSIS
    Scans profile root for profiles with FolderPattern regex matching files in the specified folder.

.PARAMETER folderPath
    Source folder path to check files from.

.OUTPUTS
    [PSCustomObject] Profile item object with Name and Path or null if no match found.
#>
function Find-MatchingProfile([string]$folderPath) {
    if (-not (Test-Path -LiteralPath $folderPath -PathType Container)) { return $null }
    $files = @(Get-ChildItem -LiteralPath $folderPath -File -ErrorAction SilentlyContinue |
               Select-Object -First 20 | ForEach-Object { $_.BaseName })
    if ($files.Count -eq 0) { return $null }

    $profileFiles = @(Get-ChildItem $script:ProfileRoot -Filter '*.json' -ErrorAction SilentlyContinue)
    foreach ($file in $profileFiles) {
        try {
            $raw = Get-Content -LiteralPath $file.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
            $pat = if ($raw.PSObject.Properties['FolderPattern']) { [string]$raw.FolderPattern } else { '' }
            if ([string]::IsNullOrWhiteSpace($pat)) { continue }
            foreach ($name in $files) {
                if ($name -match $pat) {
                    return [pscustomobject]@{ Name = $file.BaseName; Path = $file.FullName }
                }
            }
        }
        catch {}
    }
    return $null
}
