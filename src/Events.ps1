# Events.ps1 — Wiring up all WPF event handlers.
# Dot-sourced by the main script after XAML controls are bound.

# --- Folder browse ---
$BrowseSource.Add_Click({
        $p = FolderDialog (T 'Msg_SelectSource') $SourcePath.Text
        if ($p) { $SourcePath.Text = $p }
    })
$BrowseDestination.Add_Click({
        $p = FolderDialog (T 'Msg_SelectDest') $DestinationPath.Text
        if ($p) { $DestinationPath.Text = $p }
    })

# --- Tab 0: Compliance ---
if ($ComplianceScan -and $ComplianceFixSelected -and $ComplianceExtFilter) {
    $ComplianceScan.Add_Click({
            try { ScanCompliance; SetStatus (T 'Status_ComplianceDone') }
            catch { ErrorBox (T 'Err_Compliance') $_ }
        })
    $ComplianceFixSelected.Add_Click({
            try { ApplyComplianceFix }
            catch { ErrorBox (T 'Err_Compliance') $_ }
        })
    $ComplianceExtFilter.Add_SelectionChanged({
            try { ScanCompliance } catch {}
        })
}

# --- Tab 2: Analysis ---
$Analyze.Add_Click({
        try { AnalyzePatterns; SetStatus (T 'Status_AnalysisDone') }
        catch { ErrorBox (T 'Err_Analysis') $_ }
    })
$ExtensionFilter.Add_SelectionChanged({
        try { BuildPatternList } catch {}
    })
$PatternGrid.Add_SelectionChanged({
        if ($PatternGrid.SelectedItem) {
            SetPattern $PatternGrid.SelectedItem
        }
    })

# --- Tab 3: Fields ---
$FieldGrid.Add_SelectionChanged({
        $f = $FieldGrid.SelectedItem
        if ($f) {
            $FieldName.Text = $f.Name
            # Select matching role item
            $FieldRole.SelectedItem = @($FieldRole.Items | Where-Object { $_.Content -eq $f.Role })[0]
            if (-not $FieldRole.SelectedItem) { $FieldRole.SelectedIndex = 0 }
            $typeOptions = @(GetFieldTypeOptions $f)
            $FieldType.ItemsSource = $typeOptions
            $selectedType = @($typeOptions | Where-Object {
                    $_.Id -eq $f.SelectedTypeId -and [string]$_.Format -eq [string]$f.SelectedFormat
                } | Select-Object -First 1)
            if ($selectedType.Count -gt 0) { $FieldType.SelectedItem = $selectedType[0] }
            else { $FieldType.SelectedIndex = 0 }
            $CandidateInfo.Text = GetFieldCandidateSummary $f
            # Show transforms for this field
            $TransformList.ItemsSource = $null
            if ($f.Transforms) {
                $TransformList.ItemsSource = @($f.Transforms)
            }
        }
    })

$FieldApply.Add_Click({
        try {
            $f = $FieldGrid.SelectedItem
            if (-not $f) { throw (T 'Err_SelFieldTab') }
            $newName = $FieldName.Text.Trim()
            if (-not $newName) { throw (T 'Err_FieldName') }
            $oldName = $f.Name

            $f.Name = $newName
            $f.Role = [string]$FieldRole.SelectedItem.Content
            $selectedType = $FieldType.SelectedItem
            $f.SelectedTypeId = if ($selectedType) { [string]$selectedType.Id } else { 'Auto' }
            $f.SelectedFormat = if ($selectedType) { [string]$selectedType.Format } else { $null }
            $f.EffectiveType = GetEffectiveTypeLabel $f
            $f.TypeStatus = GetFieldTypeStatus $f
            $f.CandidateSummary = GetFieldCandidateSummary $f

            # Update references in OutputParts and Mappings when name changed
            if ($oldName -ne $newName) {
                foreach ($p in $script:OutputParts) {
                    if ($p.Type -eq 'Field' -and $p.Value -eq $oldName) {
                        $p.Value = $newName
                        $p.Display = "$(T 'Tag_Field') $newName"
                    }
                }
                foreach ($m in $script:Mappings) {
                    $changed = $false
                    if ($m.InputField -eq $oldName) { $m.InputField = $newName; $changed = $true }
                    if ($m.OutputField -eq $oldName) { $m.OutputField = $newName; $changed = $true }
                    if ($changed) { $m.Display = "$($m.Name): $($m.InputField) → $($m.OutputField)" }
                }
                $OutputList.ItemsSource = $null
                $OutputList.ItemsSource = $script:OutputParts
                $MappingList.ItemsSource = $null
                $MappingList.ItemsSource = $script:Mappings
            }

            $FieldGrid.ItemsSource = $null
            $FieldGrid.ItemsSource = $script:Fields
            RefreshFieldSelector
            UpdateOutputExample
        }
        catch { ErrorBox (T 'Err_Field') $_ }
    })

# --- Tab 3: Transforms ---
$TransformAdd.Add_Click({
        try {
            $f = $FieldGrid.SelectedItem
            if (-not $f) { throw (T 'Err_SelFieldTrans') }
            $result = ShowTransformDialog $f
            if ($result) {
                if (-not $f.Transforms) {
                    $f.Transforms = [System.Collections.ArrayList]::new()
                }
                [void]$f.Transforms.Add($result)
                $TransformList.ItemsSource = $null
                $TransformList.ItemsSource = @($f.Transforms)
                UpdateOutputExample
            }
        }
        catch { ErrorBox (T 'Err_Transform') $_ }
    })

$TransformRemove.Add_Click({
        $f = $FieldGrid.SelectedItem
        $i = $TransformList.SelectedIndex
        if ($f -and $f.Transforms -and $i -ge 0) {
            $f.Transforms.RemoveAt($i)
            $TransformList.ItemsSource = $null
            $TransformList.ItemsSource = @($f.Transforms)
            UpdateOutputExample
        }
    })

# --- Tab 3: Mappings ---
$MappingAdd.Add_Click({
        try { AddMappingDialog; UpdateOutputExample }
        catch { ErrorBox (T 'Err_Mapping') $_ }
    })

$MappingEdit.Add_Click({
        $i = $MappingList.SelectedIndex
        if ($i -lt 0) { throw (T 'Err_SelMapping') }
        try { AddMappingDialog $script:Mappings[$i]; UpdateOutputExample }
        catch { ErrorBox (T 'Err_Mapping') $_ }
    })

$MappingRemove.Add_Click({
        $i = $MappingList.SelectedIndex
        if ($i -ge 0) {
            # Remove virtual field if no other mapping produces it
            $removedOutput = $script:Mappings[$i].OutputField
            $script:Mappings.RemoveAt($i)

            $stillUsed = $false
            foreach ($m in $script:Mappings) {
                if ($m.OutputField -eq $removedOutput) { $stillUsed = $true; break }
            }
            if (-not $stillUsed) {
                $toRemove = $null
                foreach ($f in $script:Fields) {
                    if ($f.IsVirtual -and $f.Name -eq $removedOutput) { $toRemove = $f; break }
                }
                if ($toRemove) { $script:Fields.Remove($toRemove) }
            }

            $MappingList.ItemsSource = $null
            $MappingList.ItemsSource = $script:Mappings
            $FieldGrid.ItemsSource = $null
            $FieldGrid.ItemsSource = $script:Fields
            RefreshFieldSelector
            UpdateOutputExample
        }
    })

$MappingUp.Add_Click({
        $i = $MappingList.SelectedIndex
        if ($i -gt 0) {
            $x = $script:Mappings[$i]
            $script:Mappings.RemoveAt($i)
            $script:Mappings.Insert($i - 1, $x)
            $MappingList.ItemsSource = $null
            $MappingList.ItemsSource = $script:Mappings
            $MappingList.SelectedIndex = $i - 1
        }
    })

$MappingDown.Add_Click({
        $i = $MappingList.SelectedIndex
        if ($i -ge 0 -and $i -lt $script:Mappings.Count - 1) {
            $x = $script:Mappings[$i]
            $script:Mappings.RemoveAt($i)
            $script:Mappings.Insert($i + 1, $x)
            $MappingList.ItemsSource = $null
            $MappingList.ItemsSource = $script:Mappings
            $MappingList.SelectedIndex = $i + 1
        }
    })

# --- Tab 4: Output name builder ---
$OutputAddField.Add_Click({
        try {
            $fieldName = [string]$FieldSelector.SelectedItem
            if (-not $fieldName) { throw (T 'Err_SelFieldCombo') }
            $script:OutputParts.Add([pscustomobject]@{
                    Type    = 'Field'
                    Value   = $fieldName
                    Display = "$(T 'Tag_Field') $fieldName"
                })
            $OutputList.ItemsSource = $null
            $OutputList.ItemsSource = $script:OutputParts
            UpdateOutputExample
        }
        catch { ErrorBox (T 'Title_Error') $_ }
    })

$OutputAddText.Add_Click({
        $v = $OutputText.Text
        if ($v) {
            $script:OutputParts.Add([pscustomobject]@{
                    Type    = 'Text'
                    Value   = $v
                    Display = "$(T 'Tag_Text') $v"
                })
            $OutputText.Clear()
            $OutputList.ItemsSource = $null
            $OutputList.ItemsSource = $script:OutputParts
            UpdateOutputExample
        }
    })

$OutputAddSeparator.Add_Click({
        $script:OutputParts.Add([pscustomobject]@{
                Type    = 'Text'
                Value   = '_'
                Display = "$(T 'Tag_Separator') _"
            })
        $OutputList.ItemsSource = $null
        $OutputList.ItemsSource = $script:OutputParts
        UpdateOutputExample
    })

$OutputRemove.Add_Click({
        $i = $OutputList.SelectedIndex
        if ($i -ge 0) {
            $script:OutputParts.RemoveAt($i)
            $OutputList.ItemsSource = $null
            $OutputList.ItemsSource = $script:OutputParts
            UpdateOutputExample
        }
    })

$OutputUp.Add_Click({
        $i = $OutputList.SelectedIndex
        if ($i -gt 0) {
            $x = $script:OutputParts[$i]
            $script:OutputParts.RemoveAt($i)
            $script:OutputParts.Insert($i - 1, $x)
            $OutputList.ItemsSource = $null
            $OutputList.ItemsSource = $script:OutputParts
            $OutputList.SelectedIndex = $i - 1
            UpdateOutputExample
        }
    })

$OutputDown.Add_Click({
        $i = $OutputList.SelectedIndex
        if ($i -ge 0 -and $i -lt $script:OutputParts.Count - 1) {
            $x = $script:OutputParts[$i]
            $script:OutputParts.RemoveAt($i)
            $script:OutputParts.Insert($i + 1, $x)
            $OutputList.ItemsSource = $null
            $OutputList.ItemsSource = $script:OutputParts
            $OutputList.SelectedIndex = $i + 1
            UpdateOutputExample
        }
    })

$KeepExtension.Add_Checked({ $NewExtension.IsEnabled = $false })
$KeepExtension.Add_Unchecked({ $NewExtension.IsEnabled = $true })

# --- Tab 5: Preview and execution ---
$BuildPreview.Add_Click({
        try { FullBuildPreview; SetStatus (T 'Status_PrevBuilt') }
        catch { ErrorBox (T 'Err_Preview') $_ }
    })

$PreviewFilter.Add_SelectionChanged({ RefreshPreviewGrid })

$ExportAudit.Add_Click({
        try {
            if (-not $script:PreviewRows.Count) { throw (T 'Err_BuildPrev1') }
            $dlg = New-Object Microsoft.Win32.SaveFileDialog
            $dlg.Filter = 'CSV (*.csv)|*.csv'
            $dlg.FileName = 'FileNameTransformationAudit.csv'
            if ($dlg.ShowDialog()) {
                $script:PreviewRows | Export-Csv -LiteralPath $dlg.FileName -NoTypeInformation -Encoding UTF8 -UseCulture
                SetStatus "$(T 'Status_Exported') $($dlg.FileName)"
            }
        }
        catch { ErrorBox (T 'Err_Export') $_ }
    })

$OpenLog.Add_Click({
        Start-Process notepad.exe "`"$script:LogPath`""
    })

$Execute.Add_Click({
        try { ExecuteCopy }
        catch { ErrorBox (T 'Err_Execute') $_ }
    })

# --- Tab 1: Profiles ---
$ProfileNew.Add_Click({
        $script:Fields.Clear()
        $script:Mappings.Clear()
        $script:OutputParts.Clear()
        $FieldGrid.ItemsSource = $script:Fields
        $MappingList.ItemsSource = $script:Mappings
        $OutputList.ItemsSource = $script:OutputParts
        $TransformList.ItemsSource = $null
        RefreshFieldSelector
        $script:CurrentProfileName = ''
        $CurrentProfile.Text = (T 'Txt_Unsaved')
        UpdateOutputExample
    })

$ProfileSave.Add_Click({
        try { SaveProfile }
        catch { ErrorBox (T 'Err_SaveProfile') $_ }
    })

$ProfileLoad.Add_Click({
        try {
            if (-not $ProfileList.SelectedItem) { throw (T 'Err_SelProfList') }
            LoadProfile $ProfileList.SelectedItem.Path
        }
        catch { ErrorBox (T 'Err_LoadProfile') $_ }
    })

$ProfileCopy.Add_Click({
        try {
            if (-not $ProfileList.SelectedItem) { throw (T 'Err_SelProfCopy') }
            LoadProfile $ProfileList.SelectedItem.Path
            $script:CurrentProfileName = (T 'Prefix_Copy') + $script:CurrentProfileName
            SaveProfile
        }
        catch { ErrorBox (T 'Err_CopyProfile') $_ }
    })

$ProfileDelete.Add_Click({
        try {
            if (-not $ProfileList.SelectedItem) { throw (T 'Err_SelProfDel') }
            $name = $ProfileList.SelectedItem.Name
            $confirm = [Windows.MessageBox]::Show(
                "$(T 'Msg_DelProf') '$name'?", (T 'Title_Confirm'), 'YesNo', 'Warning'
            )
            if ($confirm -eq 'Yes') {
                Remove-Item $ProfileList.SelectedItem.Path -Force
                RefreshProfiles
                SetStatus "$(T 'Status_ProfDel') $name"
            }
        }
        catch { ErrorBox (T 'Err_DelProfile') $_ }
    })
