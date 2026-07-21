# Transforms.ps1 — Transform engine and dialog for adding per-field transformations.
# Dot-sourced by the main script; operates in $script: scope.

function ApplyTransforms([string]$value, $transforms) {
    foreach ($t in $transforms) {
        try {
            switch ($t.Type) {
                'Substring' {
                    $start = [int]$t.Start
                    $len = [int]$t.Length
                    if ($start -ge $value.Length) {
                        throw "Pozycja $start przekracza długość tekstu ($($value.Length))."
                    }
                    $len = [Math]::Min($len, $value.Length - $start)
                    $value = $value.Substring($start, $len)
                }
                'DateFormat' {
                    $parsed = [datetime]::ParseExact(
                        $value, $t.InputFormat,
                        [Globalization.CultureInfo]::InvariantCulture
                    )
                    $value = $parsed.ToString($t.OutputFormat)
                }
                'Replace' {
                    $value = $value.Replace($t.OldText, $t.NewText)
                }
                'Case' {
                    switch ($t.Mode) {
                        'Upper' { $value = $value.ToUpper() }
                        'Lower' { $value = $value.ToLower() }
                        'Title' { $value = (Get-Culture).TextInfo.ToTitleCase($value.ToLower()) }
                    }
                }
                'Pad' {
                    $ch = if ($t.PadChar) { [char]$t.PadChar[0] } else { [char]' ' }
                    $w = [int]$t.Width
                    if ($t.Side -eq 'Left') { $value = $value.PadLeft($w, $ch) }
                    else { $value = $value.PadRight($w, $ch) }
                }
                'Number' {
                    $numeric = 0.0
                    if (-not [double]::TryParse($value, [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$numeric)) {
                        throw "$(T 'Err_NotNumeric'): '$value'"
                    }
                    switch ($t.Operation) {
                        'Add' { $numeric = $numeric + $t.Value }
                        'Subtract' { $numeric = $numeric - $t.Value }
                        'Multiply' { $numeric = $numeric * $t.Value }
                        'Divide' {
                            if ($t.Value -eq 0) { throw (T 'Err_DivZero') }
                            $numeric = $numeric / $t.Value
                        }
                        'Round' { $numeric = [Math]::Round($numeric, [int]$t.Value) }
                    }
                    $value = if ($t.Format) {
                        if ($t.Format -match '^[DXdx]\d*$') {
                            # D/X format specifiers require an integral type in .NET
                            ([long][Math]::Round($numeric)).ToString($t.Format, [Globalization.CultureInfo]::InvariantCulture)
                        }
                        else {
                            $numeric.ToString($t.Format, [Globalization.CultureInfo]::InvariantCulture)
                        }
                    }
                    else {
                        $numeric.ToString([Globalization.CultureInfo]::InvariantCulture)
                    }
                }
            }
        }
        catch {
            throw "$(T 'Err_Transform') '$($t.Display)': $($_.Exception.Message)"
        }
    }
    return $value
}

function ShowTransformDialog($field) {
    $form = New-Object Windows.Forms.Form
    $form.Text = (T 'Title_AddTransform')
    $form.Size = New-Object Drawing.Size(520, 300)
    $form.StartPosition = 'CenterParent'
    $form.Font = New-Object Drawing.Font('Segoe UI', 9)
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $defaultDateInputFormat = 'yyyyMMdd'
    if ($field) {
        $resolvedType = GetResolvedFieldType $field
        if ($resolvedType.TypeId -eq 'DateTime' -and $resolvedType.Format) {
            $defaultDateInputFormat = [string]$resolvedType.Format
        }
    }

    # Type label and combobox
    $lblType = New-Object Windows.Forms.Label
    $lblType.Text = (T 'Txt_TransformType')
    $lblType.Location = New-Object Drawing.Point(15, 18)
    $lblType.Size = New-Object Drawing.Size(140, 22)
    $form.Controls.Add($lblType)

    $cboType = New-Object Windows.Forms.ComboBox
    $cboType.DropDownStyle = 'DropDownList'
    $cboType.Location = New-Object Drawing.Point(160, 15)
    $cboType.Size = New-Object Drawing.Size(320, 25)
    $cboType.Items.AddRange(@(
            (T 'Tr_Substring'),
            (T 'Tr_DateFormat'),
            (T 'Tr_Replace'),
            (T 'Tr_Case'),
            (T 'Tr_Pad')
        ))

    # Number/Math transform is only applicable to numeric field types
    $effectiveTypeId = if ($field) {
        $resolvedType = GetResolvedFieldType $field
        [string]$resolvedType.TypeId
    } else { '' }
    $isMathField = $effectiveTypeId -in @('Integer', 'Decimal')
    if ($isMathField) {
        [void]$cboType.Items.Add((T 'Tr_Number'))
    }
    $form.Controls.Add($cboType)

    # 3 parameter rows: label + textbox
    $paramLabels = @()
    $paramControls = @()
    for ($row = 0; $row -lt 3; $row++) {
        $lbl = New-Object Windows.Forms.Label
        $lbl.Location = New-Object Drawing.Point(15, (60 + $row * 38))
        $lbl.Size = New-Object Drawing.Size(140, 22)
        $lbl.Visible = $false
        $form.Controls.Add($lbl)
        $paramLabels += $lbl

        $txt = New-Object Windows.Forms.TextBox
        $txt.Location = New-Object Drawing.Point(160, (57 + $row * 38))
        $txt.Size = New-Object Drawing.Size(320, 25)
        $txt.Visible = $false
        $form.Controls.Add($txt)
        $paramControls += $txt
    }

    # Alternative ComboBox for param1 (Case mode, Pad side)
    $cboParam1 = New-Object Windows.Forms.ComboBox
    $cboParam1.DropDownStyle = 'DropDownList'
    $cboParam1.Location = New-Object Drawing.Point(160, 57)
    $cboParam1.Size = New-Object Drawing.Size(320, 25)
    $cboParam1.Visible = $false
    $form.Controls.Add($cboParam1)

    # Update visible parameters when type changes
    $cboType.Add_SelectedIndexChanged({
            foreach ($l in $paramLabels) { $l.Visible = $false }
            foreach ($c in $paramControls) { $c.Visible = $false; $c.Text = '' }
            $cboParam1.Visible = $false
            $cboParam1.Items.Clear()

            switch ($cboType.SelectedIndex) {
                0 {
                    # Substring
                    $paramLabels[0].Text = (T 'Tr_PosStart'); $paramLabels[0].Visible = $true
                    $paramControls[0].Text = '0'; $paramControls[0].Visible = $true
                    $paramLabels[1].Text = (T 'Tr_CharCount'); $paramLabels[1].Visible = $true
                    $paramControls[1].Visible = $true
                }
                1 {
                    # DateFormat
                    $paramLabels[0].Text = (T 'Tr_FmtIn'); $paramLabels[0].Visible = $true
                    $paramControls[0].Text = $defaultDateInputFormat; $paramControls[0].Visible = $true
                    $paramLabels[1].Text = (T 'Tr_FmtOut'); $paramLabels[1].Visible = $true
                    $paramControls[1].Visible = $true
                }
                2 {
                    # Replace
                    $paramLabels[0].Text = (T 'Tr_Search'); $paramLabels[0].Visible = $true
                    $paramControls[0].Visible = $true
                    $paramLabels[1].Text = (T 'Tr_NewTxt'); $paramLabels[1].Visible = $true
                    $paramControls[1].Visible = $true
                }
                3 {
                    # Case
                    $paramLabels[0].Text = (T 'Tr_Mode'); $paramLabels[0].Visible = $true
                    $cboParam1.Items.AddRange(@((T 'Tr_Upper'), (T 'Tr_Lower'), (T 'Tr_Title')))
                    $cboParam1.SelectedIndex = 0
                    $cboParam1.Visible = $true
                }
                4 {
                    # Pad
                    $paramLabels[0].Text = (T 'Tr_Side'); $paramLabels[0].Visible = $true
                    $cboParam1.Items.AddRange(@((T 'Tr_Left'), (T 'Tr_Right')))
                    $cboParam1.SelectedIndex = 0
                    $cboParam1.Visible = $true
                    $paramLabels[1].Text = (T 'Tr_PadChar'); $paramLabels[1].Visible = $true
                    $paramControls[1].Text = '0'; $paramControls[1].Visible = $true
                    $paramLabels[2].Text = (T 'Tr_TargetLen'); $paramLabels[2].Visible = $true
                    $paramControls[2].Visible = $true
                }
                5 {
                    # Number / Math
                    $paramLabels[0].Text = (T 'Tr_Operation'); $paramLabels[0].Visible = $true
                    $cboParam1.Items.AddRange(@(
                            (T 'Tr_OpAdd'), (T 'Tr_OpSubtract'), (T 'Tr_OpMultiply'), (T 'Tr_OpDivide'), (T 'Tr_OpRound')
                        ))
                    $cboParam1.SelectedIndex = 0
                    $cboParam1.Visible = $true
                    $paramLabels[1].Text = (T 'Tr_OpValue'); $paramLabels[1].Visible = $true
                    $paramControls[1].Text = '0'; $paramControls[1].Visible = $true
                    $paramLabels[2].Text = (T 'Tr_NumFormat'); $paramLabels[2].Visible = $true
                    $paramControls[2].Visible = $true
                }
            }
        })

    # Add button
    $btnAdd = New-Object Windows.Forms.Button
    $btnAdd.Text = (T 'Btn_Add')
    $btnAdd.Location = New-Object Drawing.Point(400, 220)
    $btnAdd.Size = New-Object Drawing.Size(80, 30)
    $form.Controls.Add($btnAdd)

    $script:transformDialogResult = $null

    $btnAdd.Add_Click({
            try {
                $result = switch ($cboType.SelectedIndex) {
                    0 {
                        # Substring
                        $s = [int]$paramControls[0].Text
                        $l = [int]$paramControls[1].Text
                        if ($l -le 0) { throw (T 'Err_CharCount') }
                        [pscustomobject]@{
                            Type = 'Substring'; Start = $s; Length = $l
                            Display = "$(T 'Disp_Substr') $s, $l $(T 'Disp_Chars')"
                        }
                    }
                    1 {
                        # DateFormat
                        $inf = $paramControls[0].Text.Trim()
                        $outf = $paramControls[1].Text.Trim()
                        if (-not $inf -or -not $outf) { throw (T 'Err_BothFmt') }
                        [pscustomobject]@{
                            Type = 'DateFormat'; InputFormat = $inf; OutputFormat = $outf
                            Display = "$(T 'Disp_Date') $inf → $outf"
                        }
                    }
                    2 {
                        # Replace
                        $old = $paramControls[0].Text
                        $new = $paramControls[1].Text
                        if ($old -eq '') { throw (T 'Err_SearchTxt') }
                        [pscustomobject]@{
                            Type = 'Replace'; OldText = $old; NewText = $new
                            Display = "$(T 'Disp_Replace') '$old' → '$new'"
                        }
                    }
                    3 {
                        # Case
                        if ($cboParam1.SelectedIndex -lt 0) { throw (T 'Err_SelMode') }
                        $mode = @('Upper', 'Lower', 'Title')[$cboParam1.SelectedIndex]
                        $modeText = [string]$cboParam1.SelectedItem
                        [pscustomobject]@{
                            Type = 'Case'; Mode = $mode
                            Display = "$(T 'Disp_Case') $modeText"
                        }
                    }
                    4 {
                        # Pad
                        if ($cboParam1.SelectedIndex -lt 0) { throw (T 'Err_SelSide') }
                        $side = @('Left', 'Right')[$cboParam1.SelectedIndex]
                        $ch = $paramControls[1].Text
                        $width = [int]$paramControls[2].Text
                        if (-not $ch) { throw (T 'Err_PadChar') }
                        if ($width -le 0) { throw (T 'Err_TargetLen') }
                        $sideText = [string]$cboParam1.SelectedItem
                        [pscustomobject]@{
                            Type = 'Pad'; Side = $side; PadChar = $ch; Width = $width
                            Display = "$(T 'Disp_Pad') '$ch' $sideText $(T 'Disp_To') $width $(T 'Disp_Chars')"
                        }
                    }
                    5 {
                        # Number / Math
                        if ($cboParam1.SelectedIndex -lt 0) { throw (T 'Err_SelOperation') }
                        $operation = @('Add', 'Subtract', 'Multiply', 'Divide', 'Round')[$cboParam1.SelectedIndex]
                        $operationText = [string]$cboParam1.SelectedItem
                        $valueText = $paramControls[1].Text.Trim()
                        $numValue = 0.0
                        if (-not [double]::TryParse($valueText, [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$numValue)) {
                            throw (T 'Err_NumValue')
                        }
                        if ($operation -eq 'Divide' -and $numValue -eq 0) { throw (T 'Err_DivZero') }
                        $numFormat = $paramControls[2].Text.Trim()
                        $displaySuffix = if ($numFormat) { " ($numFormat)" } else { '' }
                        [pscustomobject]@{
                            Type = 'Number'; Operation = $operation; Value = $numValue; Format = $numFormat
                            Display = "$(T 'Disp_Number') $operationText $numValue$displaySuffix"
                        }
                    }
                    default { throw (T 'Err_SelTransform') }
                }
                $script:transformDialogResult = $result
                $form.Close()
            }
            catch {
                [Windows.Forms.MessageBox]::Show($_.Exception.Message, (T 'Title_Error'), 'OK', 'Warning') | Out-Null
            }
        })

    $cboType.SelectedIndex = 0
    [void]$form.ShowDialog()
    return $script:transformDialogResult
}
