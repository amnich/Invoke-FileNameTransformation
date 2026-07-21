# Profiles.ps1 — Saving, loading, and refreshing JSON profiles.
# Dot-sourced by the main script; operates in $script: scope.

function RefreshProfiles {
    $items = @(
        Get-ChildItem $script:ProfileRoot -Filter '*.json' -ErrorAction SilentlyContinue |
        ForEach-Object { [pscustomobject]@{ Name = $_.BaseName; Path = $_.FullName } }
    )
    $ProfileList.ItemsSource = $items
}

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
    }
    $obj = ConvertTo-FNTProfile ([pscustomobject]$obj)

    $path = Join-Path $script:ProfileRoot ($name + '.json')
    $obj | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $path -Encoding UTF8

    $script:CurrentProfileName = $name
    $CurrentProfile.Text = $name
    RefreshProfiles
    SetStatus "$(T 'Status_ProfSaved') $name"
}

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
