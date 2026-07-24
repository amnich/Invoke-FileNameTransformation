# Mappings.ps1 — CSV mapping functions and the mapping dialog.
# Dot-sourced by the main script; operates in $script: scope.

<#
.SYNOPSIS
    Reads column headers and detects delimiters from CSV, TXT, JSON, or XML mapping files.

.PARAMETER path
    File path to inspect.

.OUTPUTS
    [PSCustomObject] Object containing Delimiter, Headers array, and Format ('CSV', 'JSON', 'XML').
#>
function Get-FNTDictionaryHeaders([string]$path) {
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        return [pscustomobject]@{ Delimiter = ','; Headers = @(); Format = 'CSV' }
    }
    $ext = [System.IO.Path]::GetExtension($path).ToLowerInvariant()

    if ($ext -eq '.json') {
        try {
            $raw = Get-Content -LiteralPath $path -Raw -Encoding UTF8
            $json = $raw | ConvertFrom-Json
            $headers = @()
            if ($json -is [System.Array] -and $json.Count -gt 0) {
                $headers = @($json[0].PSObject.Properties.Name)
            }
            elseif ($json -is [PSCustomObject] -or $json -is [hashtable]) {
                $headers = @($json.PSObject.Properties.Name)
            }
            return [pscustomobject]@{ Delimiter = ''; Headers = $headers; Format = 'JSON' }
        }
        catch {
            return [pscustomobject]@{ Delimiter = ''; Headers = @(); Format = 'JSON' }
        }
    }
    elseif ($ext -eq '.xml') {
        try {
            [xml]$xml = Get-Content -LiteralPath $path -Raw -Encoding UTF8
            $headers = New-Object System.Collections.Generic.List[string]
            $root = $xml.DocumentElement
            if ($root.HasChildNodes) {
                $firstChild = $root.ChildNodes[0]
                foreach ($node in $firstChild.ChildNodes) {
                    if ($node.Name) { $headers.Add([string]$node.Name) }
                }
                foreach ($attr in $firstChild.Attributes) {
                    if ($attr.Name) { $headers.Add("@" + $attr.Name) }
                }
            }
            return [pscustomobject]@{ Delimiter = ''; Headers = @($headers | Select-Object -Unique); Format = 'XML' }
        }
        catch {
            return [pscustomobject]@{ Delimiter = ''; Headers = @(); Format = 'XML' }
        }
    }
    else {
        # CSV auto-detection
        $headerLine = Get-Content -LiteralPath $path -Encoding UTF8 | Where-Object { $_ -and $_.Trim() } | Select-Object -First 1
        if (-not $headerLine) {
            return [pscustomobject]@{ Delimiter = ','; Headers = @(); Format = 'CSV' }
        }
        $delimiter = @(';', ',', "`t", '|') |
        Sort-Object { ([regex]::Matches($headerLine, [regex]::Escape($_))).Count } -Descending |
        Select-Object -First 1

        $headers = @()
        if ($headerLine) {
            $headers = @($headerLine -split [regex]::Escape($delimiter))
            $headers = @($headers | ForEach-Object { $_.Trim().Trim('"') } | Where-Object { $_ -ne '' })
        }
        return [pscustomobject]@{ Delimiter = $delimiter; Headers = $headers; Format = 'CSV' }
    }
}

<#
.SYNOPSIS
    Alias for Get-FNTDictionaryHeaders.
#>
function CsvHeaders([string]$path) {
    return Get-FNTDictionaryHeaders $path
}

<#
.SYNOPSIS
    Ensures a virtual field entry exists in $script:Fields for mapping or metadata outputs.

.PARAMETER name
    The business field name for the virtual field.
#>
function EnsureVirtualField([string]$name) {
    $exists = $false
    foreach ($f in $script:Fields) {
        if ($f.Name -eq $name) { $exists = $true; break }
    }
    if (-not $exists) {
        $script:Fields.Add([pscustomobject]@{
                PartIndex    = -1
                DisplayIndex = 'V'
                Sample       = (T 'Val_Mapping')
                Preview      = ''
                DetectedType = (T 'Src_Mapping')
                Name         = $name
                Role         = (T 'Role_Value')
                IsVirtual    = $true
                Source       = (T 'Src_Mapping')
                Transforms   = [System.Collections.ArrayList]::new()
            })
    }
}

<#
.SYNOPSIS
    Opens a modal Windows Forms dialog to create or edit a dictionary mapping rule.

.PARAMETER mapping
    Optional existing mapping object to edit. If null, creates a new mapping.
#>
function AddMappingDialog([object]$mapping = $null) {
    $form = New-Object Windows.Forms.Form
    $form.Text = if ($mapping) { (T 'Title_EditMapping') } else { (T 'Title_AddMapping') }
    $form.Size = New-Object Drawing.Size(650, 310)
    $form.StartPosition = 'CenterParent'
    $form.Font = New-Object Drawing.Font('Segoe UI', 9)
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false

    $labels = @(
        (T 'Lbl_MapName'),
        (T 'Lbl_MapIn'),
        (T 'Lbl_MapOut'),
        (T 'Lbl_MapFile'),
        (T 'Lbl_MapKey'),
        (T 'Lbl_MapVal')
    )
    $controls = @()

    for ($i = 0; $i -lt 6; $i++) {
        $lbl = New-Object Windows.Forms.Label
        $lbl.Text = $labels[$i]
        $lbl.Location = New-Object Drawing.Point(15, (15 + $i * 35))
        $lbl.Size = New-Object Drawing.Size(140, 25)
        $form.Controls.Add($lbl)

        if ($i -in 1, 4, 5) {
            $ctl = New-Object Windows.Forms.ComboBox
            $ctl.DropDownStyle = 'DropDownList'
        }
        else {
            $ctl = New-Object Windows.Forms.TextBox
        }
        $ctl.Location = New-Object Drawing.Point(160, (12 + $i * 35))
        $ctl.Size = New-Object Drawing.Size(350, 25)
        $form.Controls.Add($ctl)
        $controls += $ctl
    }

    # Populate input field combo with all field names
    $controls[1].Items.AddRange(@($script:Fields | ForEach-Object { $_.Name }))

    if ($mapping) {
        $controls[0].Text = [string]$mapping.Name
        $controls[1].SelectedItem = $mapping.InputField
        $controls[2].Text = [string]$mapping.OutputField
        $controls[3].Text = [string]$mapping.Path
        $controls[4].SelectedItem = $mapping.KeyColumn
        $controls[5].SelectedItem = $mapping.ValueColumn
        if ($mapping.Path) {
            try {
                $h = CsvHeaders $mapping.Path
                $script:tempDelimiter = $h.Delimiter
                $controls[4].Items.Clear()
                $controls[5].Items.Clear()
                $controls[4].Items.AddRange($h.Headers)
                $controls[5].Items.AddRange($h.Headers)
            }
            catch {
                $script:tempDelimiter = $mapping.Delimiter
            }
        }
        else {
            $script:tempDelimiter = $mapping.Delimiter
        }
    }

    $script:tempDelimiter = if ($mapping -and $mapping.Delimiter) { [string]$mapping.Delimiter } else { ',' }

    # Browse button for data file
    $browse = New-Object Windows.Forms.Button
    $browse.Text = $(T 'Btn_Browse')
    $browse.Location = New-Object Drawing.Point(520, 117)
    $browse.Size = New-Object Drawing.Size(80, 25)
    $form.Controls.Add($browse)

    $browse.Add_Click({
            $p = FileDialog
            if ($p) {
                $controls[3].Text = $p
                try {
                    $h = CsvHeaders $p
                    $script:tempDelimiter = $h.Delimiter
                    $controls[4].Items.Clear()
                    $controls[5].Items.Clear()
                    $controls[4].Items.AddRange($h.Headers)
                    $controls[5].Items.AddRange($h.Headers)
                }
                catch {
                    [Windows.Forms.MessageBox]::Show(
                        "$(T 'Err_ReadHeaders') $($_.Exception.Message)",
                        (T 'Title_FileErr'), 'OK', 'Warning'
                    ) | Out-Null
                }
            }
        })

    # Add button
    $btnOk = New-Object Windows.Forms.Button
    $btnOk.Text = if ($mapping) { (T 'Btn_Edit') } else { (T 'Btn_Add') }
    $btnOk.Location = New-Object Drawing.Point(520, 225)
    $btnOk.Size = New-Object Drawing.Size(80, 30)
    $form.Controls.Add($btnOk)

    $btnOk.Add_Click({
            if (-not $controls[0].Text -or
                $controls[1].SelectedItem -eq $null -or
                -not $controls[2].Text -or
                -not $controls[3].Text -or
                $controls[4].SelectedItem -eq $null -or
                $controls[5].SelectedItem -eq $null) {
                [Windows.Forms.MessageBox]::Show((T 'Err_FillFields'), (T 'Title_MissingData'), 'OK', 'Warning') | Out-Null
                return
            }

            $targetMapping = $mapping
            if ($targetMapping) {
                $oldOutputField = [string]$targetMapping.OutputField
                $targetMapping.Name = $controls[0].Text
                $targetMapping.InputField = [string]$controls[1].SelectedItem
                $targetMapping.OutputField = $controls[2].Text.Trim()
                $targetMapping.Path = $controls[3].Text
                $targetMapping.KeyColumn = [string]$controls[4].SelectedItem
                $targetMapping.ValueColumn = [string]$controls[5].SelectedItem
                $targetMapping.Delimiter = $script:tempDelimiter
                $targetMapping.Display = "$($controls[0].Text): $($controls[1].SelectedItem) → $($controls[2].Text)"
            }
            else {
                $targetMapping = [pscustomobject]@{
                    Name        = $controls[0].Text
                    InputField  = [string]$controls[1].SelectedItem
                    OutputField = $controls[2].Text.Trim()
                    Path        = $controls[3].Text
                    KeyColumn   = [string]$controls[4].SelectedItem
                    ValueColumn = [string]$controls[5].SelectedItem
                    Delimiter   = $script:tempDelimiter
                    Display     = "$($controls[0].Text): $($controls[1].SelectedItem) → $($controls[2].Text)"
                }
                $script:Mappings.Add($targetMapping)
            }

            # Create or refresh virtual field for the output
            $newOutputField = [string]$targetMapping.OutputField
            if ($mapping) {
                $oldOutputField = [string]$mapping.OutputField
                if ($oldOutputField -and $oldOutputField -ne $newOutputField) {
                    $stillUsed = $false
                    foreach ($m in $script:Mappings) {
                        if ($m.OutputField -eq $oldOutputField) { $stillUsed = $true; break }
                    }
                    if (-not $stillUsed) {
                        $toRemove = $null
                        foreach ($f in $script:Fields) {
                            if ($f.IsVirtual -and $f.Name -eq $oldOutputField) { $toRemove = $f; break }
                        }
                        if ($toRemove) { $script:Fields.Remove($toRemove) }
                    }
                }
            }
            EnsureVirtualField $newOutputField

            # Refresh UI
            $MappingList.ItemsSource = $null
            $MappingList.ItemsSource = $script:Mappings
            $FieldGrid.ItemsSource = $null
            $FieldGrid.ItemsSource = $script:Fields
            RefreshFieldSelector

            $form.Close()
        })

    [void]$form.ShowDialog()
}
