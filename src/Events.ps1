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

# --- Folder path suggestions & Auto profile matching ---
$SourcePath.Add_TextChanged({
        Update-PathSuggestions $SourcePath $SourcePathSuggestions $SourcePathSuggestionList
        $folder = $SourcePath.Text.Trim()
        if ($AutoProfileHint -and $AutoProfileApply -and
            -not [string]::IsNullOrWhiteSpace($folder) -and
            (Test-Path -LiteralPath $folder -PathType Container)) {
            $matched = Find-MatchingProfile $folder
            if ($matched) {
                $msgFormat = T 'Txt_AutoProfileHint'
                $AutoProfileHint.Text = ($msgFormat -f $matched.Name)
                $AutoProfileHint.Visibility  = 'Visible'
                $AutoProfileApply.Visibility = 'Visible'
                $script:AutoMatchedProfilePath = $matched.Path
            } else {
                $AutoProfileHint.Visibility  = 'Collapsed'
                $AutoProfileApply.Visibility = 'Collapsed'
                $script:AutoMatchedProfilePath = $null
            }
        }
        elseif ($AutoProfileHint -and $AutoProfileApply) {
            $AutoProfileHint.Visibility  = 'Collapsed'
            $AutoProfileApply.Visibility = 'Collapsed'
            $script:AutoMatchedProfilePath = $null
        }
    })

if ($AutoProfileApply) {
    $AutoProfileApply.Add_Click({
            if ($script:AutoMatchedProfilePath) {
                try {
                    LoadProfile $script:AutoMatchedProfilePath
                    $AutoProfileHint.Visibility  = 'Collapsed'
                    $AutoProfileApply.Visibility = 'Collapsed'
                }
                catch { ErrorBox (T 'Err_LoadProfile') $_ }
            }
        })
}
$DestinationPath.Add_TextChanged({ Update-PathSuggestions $DestinationPath $DestinationPathSuggestions $DestinationPathSuggestionList })
$SourcePath.Add_GotKeyboardFocus({ Update-PathSuggestions $SourcePath $SourcePathSuggestions $SourcePathSuggestionList })
$DestinationPath.Add_GotKeyboardFocus({ Update-PathSuggestions $DestinationPath $DestinationPathSuggestions $DestinationPathSuggestionList })
$SourcePath.Add_PreviewKeyDown({
        if ($_.Key -eq [System.Windows.Input.Key]::Tab -and $SourcePathSuggestions.IsOpen -and $SourcePathSuggestionList.Items.Count -gt 0) {
            $SourcePathSuggestionList.SelectedIndex = 0
            Apply-PathSuggestion $SourcePath $SourcePathSuggestions $SourcePathSuggestionList
            $_.Handled = $true
        }
    })
$DestinationPath.Add_PreviewKeyDown({
        if ($_.Key -eq [System.Windows.Input.Key]::Tab -and $DestinationPathSuggestions.IsOpen -and $DestinationPathSuggestionList.Items.Count -gt 0) {
            $DestinationPathSuggestionList.SelectedIndex = 0
            Apply-PathSuggestion $DestinationPath $DestinationPathSuggestions $DestinationPathSuggestionList
            $_.Handled = $true
        }
    })
$SourcePathSuggestionList.Add_SelectionChanged({ Apply-PathSuggestion $SourcePath $SourcePathSuggestions $SourcePathSuggestionList })
$DestinationPathSuggestionList.Add_SelectionChanged({ Apply-PathSuggestion $DestinationPath $DestinationPathSuggestions $DestinationPathSuggestionList })

# --- Tab 0: Compliance ---
if ($ComplianceScan -and $ComplianceFixSelected -and $ComplianceExtFilter) {
    function Invoke-ComplianceScan {
        $previousCursor = $window.Cursor
        try {
            $window.Cursor = [Windows.Input.Cursors]::Wait
            UpdateUI
            ScanCompliance
        }
        finally {
            $window.Cursor = $previousCursor
        }
    }

    $ComplianceScan.Add_Click({
            try { Invoke-ComplianceScan; SetStatus (T 'Status_ComplianceDone') }
            catch { ErrorBox (T 'Err_Compliance') $_ }
        })
    $ComplianceFixSelected.Add_Click({
            try { ApplyComplianceFix }
            catch { ErrorBox (T 'Err_Compliance') $_ }
        })
    $ComplianceExtFilter.Add_SelectionChanged({
            try { Invoke-ComplianceScan } catch {}
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
            UpdateOutputExample
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
            UpdateOutputExample
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

if ($CollisionPolicySelector) {
    $CollisionPolicySelector.Add_SelectionChanged({
            if ($script:PreviewRows.Count -gt 0) {
                try { FullBuildPreview } catch {}
            }
        })
}

if ($Btn_UndoLastOperation) {
    $Btn_UndoLastOperation.Add_Click({
            try { Invoke-FNTUndoLastOperation }
            catch { ErrorBox (T 'Title_Error') $_ }
        })
}

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

# --- Tab 6: Custom Rules Management & Live Tester ---
if ($CustomRulesGrid) {
    function RefreshCustomRulesGrid {
        $CustomRulesGrid.ItemsSource = $null
        $safeRules = if ($script:CustomTypeRules -and $script:CustomTypeRules.Count -gt 0) {
            @($script:CustomTypeRules | Where-Object { $null -ne $_ })
        } else {
            @()
        }
        $CustomRulesGrid.ItemsSource = $safeRules
    }

    RefreshCustomRulesGrid

    $CustomRulesGrid.Add_SelectionChanged({
            $rule = $CustomRulesGrid.SelectedItem
            if ($rule) {
                $CustomRuleId.Text = [string]$rule.Id
                $CustomRuleDisplayName.Text = [string]$rule.DisplayName
                $CustomRulePattern.Text = [string]$rule.Pattern
                $CustomRuleEnabled.IsChecked = [bool]$rule.Enabled
                $CustomRuleAllowComposite.IsChecked = [bool]$rule.AllowComposite
            }
        })

    if ($CustomRuleNew) {
        $CustomRuleNew.Add_Click({
                $CustomRuleId.Text = "CustomRule$($script:CustomTypeRules.Count + 1)"
                $CustomRuleDisplayName.Text = "New Custom Rule"
                $CustomRulePattern.Text = "^[A-Z]{3}-\d{3}$"
                $CustomRuleEnabled.IsChecked = $true
                $CustomRuleAllowComposite.IsChecked = $false
            })
    }

    if ($CustomRuleApply) {
        $CustomRuleApply.Add_Click({
                try {
                    $id = $CustomRuleId.Text.Trim()
                    $displayName = $CustomRuleDisplayName.Text.Trim()
                    $pattern = $CustomRulePattern.Text.Trim()
                    $enabled = [bool]$CustomRuleEnabled.IsChecked
                    $allowComposite = [bool]$CustomRuleAllowComposite.IsChecked

                    if (-not $id -or -not $pattern) {
                        throw (T 'Err_EmptyRuleOrPattern')
                    }
                    [void][regex]::new($pattern)

                    $existing = @($script:CustomTypeRules | Where-Object { $_.Id -eq $id })[0]
                    if ($existing) {
                        $existing.DisplayName = $displayName
                        $existing.Pattern = $pattern
                        $existing.Enabled = $enabled
                        $existing.AllowComposite = $allowComposite
                    }
                    else {
                        $newRule = [pscustomobject][ordered]@{
                            Id             = $id
                            DisplayName    = $displayName
                            Pattern        = $pattern
                            Enabled        = $enabled
                            AllowComposite = $allowComposite
                        }
                        $script:CustomTypeRules += $newRule
                    }
                    RefreshCustomRulesGrid
                    SetStatus "Updated custom rule '$id'."
                }
                catch { ErrorBox "Custom Rule Error" $_ }
            })
    }

    if ($CustomRuleDelete) {
        $CustomRuleDelete.Add_Click({
                try {
                    $rule = $CustomRulesGrid.SelectedItem
                    if ($rule) {
                        $script:CustomTypeRules = @($script:CustomTypeRules | Where-Object { $_.Id -ne $rule.Id })
                        RefreshCustomRulesGrid
                        SetStatus "Deleted custom rule '$($rule.Id)'."
                    }
                }
                catch { ErrorBox "Custom Rule Error" $_ }
            })
    }

    if ($CustomRuleSaveConfig) {
        $CustomRuleSaveConfig.Add_Click({
                try {
                    $script:Config.CustomTypeRules = @($script:CustomTypeRules)
                    $script:Config | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $script:ConfigPath -Encoding UTF8
                    SetStatus "Saved custom rules configuration to $script:ConfigPath."
                }
                catch { ErrorBox "Config Save Error" $_ }
            })
    }

    if ($CustomRuleTestBtn) {
        $CustomRuleTestBtn.Add_Click({
                try {
                    $testVal = $CustomRuleTestInput.Text
                    $candidates = Get-FNTTypeCandidates -Value $testVal -CustomTypeRules @($script:CustomTypeRules)
                    if ($candidates.Count -eq 0) {
                        $CustomRuleTestResult.Text = "Input: '$testVal'`nResult: NO MATCH"
                    }
                    else {
                        $details = ($candidates | ForEach-Object { "$($_.TypeId) [Display: $($_.DisplayName), Composite: $($_.AllowComposite)]" }) -join "`n  "
                        $CustomRuleTestResult.Text = "Input: '$testVal'`nMatched Candidates ($($candidates.Count)):`n  $details"
                    }
                }
                catch {
                    $CustomRuleTestResult.Text = "Test Error: $($_.Exception.Message)"
                }
            })
    }

    if ($OpenCustomRulesDoc) {
        $OpenCustomRulesDoc.Add_Click({
                try {
                    $docPath = Join-Path $script:ScriptRoot 'CUSTOM_TYPE_RULES.md'
                    if (Test-Path -LiteralPath $docPath) {
                        [System.Diagnostics.Process]::Start($docPath) | Out-Null
                    }
                    else {
                        throw ((T 'Err_FileNotFound') -f $docPath)
                    }
                }
                catch { ErrorBox "Guide Error" $_ }
            })
    }
}
