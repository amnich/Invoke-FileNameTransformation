# Transforms.ps1 — Transform engine and dialog for adding per-field transformations.
# Dot-sourced by the main script; operates in $script: scope.

function ConvertTo-FNTAscii([string]$inputString) {
    if ([string]::IsNullOrEmpty($inputString)) { return '' }

    $map = [System.Collections.Generic.Dictionary[string, string]]::new([System.StringComparer]::Ordinal)
    # Polish
    $map['ą'] = 'a'; $map['ć'] = 'c'; $map['ę'] = 'e'; $map['ł'] = 'l'; $map['ń'] = 'n'; $map['ó'] = 'o'; $map['ś'] = 's'; $map['ź'] = 'z'; $map['ż'] = 'z'
    $map['Ą'] = 'A'; $map['Ć'] = 'C'; $map['Ę'] = 'E'; $map['Ł'] = 'L'; $map['Ń'] = 'N'; $map['Ó'] = 'O'; $map['Ś'] = 'S'; $map['Ź'] = 'Z'; $map['Ż'] = 'Z'
    # German
    $map['ä'] = 'ae'; $map['ö'] = 'oe'; $map['ü'] = 'ue'; $map['ß'] = 'ss'
    $map['Ä'] = 'Ae'; $map['Ö'] = 'Oe'; $map['Ü'] = 'Ue'
    # French
    $map['é'] = 'e'; $map['è'] = 'e'; $map['ê'] = 'e'; $map['ë'] = 'e'
    $map['É'] = 'E'; $map['È'] = 'E'; $map['Ê'] = 'E'; $map['Ë'] = 'E'
    $map['à'] = 'a'; $map['â'] = 'a'; $map['À'] = 'A'; $map['Â'] = 'A'
    $map['ç'] = 'c'; $map['Ç'] = 'C'
    # Spanish
    $map['ñ'] = 'n'; $map['Ñ'] = 'N'

    $sb = New-Object System.Text.StringBuilder
    foreach ($char in $inputString.ToCharArray()) {
        $strChar = [string]$char
        if ($map.ContainsKey($strChar)) {
            [void]$sb.Append($map[$strChar])
        }
        else {
            [void]$sb.Append($char)
        }
    }

    $normalized = $sb.ToString().Normalize([System.Text.NormalizationForm]::FormD)
    $cleanSb = New-Object System.Text.StringBuilder
    foreach ($c in $normalized.ToCharArray()) {
        $uc = [System.Globalization.CharUnicodeInfo]::GetUnicodeCategory($c)
        if ($uc -ne [System.Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$cleanSb.Append($c)
        }
    }
    return $cleanSb.ToString().Normalize([System.Text.NormalizationForm]::FormC)
}

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
                'Transliterate' {
                    $value = ConvertTo-FNTAscii $value
                }
                'RegexReplace' {
                    $value = [regex]::Replace($value, $t.Pattern, $t.Replacement)
                }
                'Expression' {
                    if ($t.Expression) {
                        $sb = [scriptblock]::Create($t.Expression)
                        $value = [string]($sb.InvokeWithContext($null, @([psvariable]::new('_', $value)), @($value)))
                    }
                }
                'Sequence' {
                    $seqVal = if ($null -ne $t.CurrentIndex) { [int]$t.CurrentIndex } else { [int]$t.Start }
                    $width = if ($t.Width) { [int]$t.Width } else { 1 }
                    $value = $seqVal.ToString().PadLeft($width, '0')
                    $t.CurrentIndex = $seqVal + (if ($t.Step) { [int]$t.Step } else { 1 })
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
    $form.Size = New-Object Drawing.Size(520, 320)
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

    $lblType = New-Object Windows.Forms.Label
    $lblType.Text = (T 'Txt_TransformType')
    $lblType.Location = New-Object Drawing.Point(15, 18)
    $lblType.Size = New-Object Drawing.Size(140, 22)
    $form.Controls.Add($lblType)

    $cboType = New-Object Windows.Forms.ComboBox
    $cboType.DropDownStyle = 'DropDownList'
    $cboType.Location = New-Object Drawing.Point(160, 15)
    $cboType.Size = New-Object Drawing.Size(320, 25)

    $transformTypeKeys = [ordered]@{
        'Substring'     = (T 'Tr_Substring')
        'DateFormat'    = (T 'Tr_DateFormat')
        'Replace'       = (T 'Tr_Replace')
        'Transliterate' = (T 'Tr_Transliterate')
        'RegexReplace'  = (T 'Tr_RegexReplace')
        'Expression'    = (T 'Tr_Expression')
        'Sequence'      = (T 'Tr_Sequence')
        'Case'          = (T 'Tr_Case')
        'Pad'           = (T 'Tr_Pad')
    }

    $effectiveTypeId = if ($field) {
        $resolvedType = GetResolvedFieldType $field
        [string]$resolvedType.TypeId
    }
    else { '' }
    $isMathField = $effectiveTypeId -in @('Integer', 'Decimal')
    if ($isMathField) {
        $transformTypeKeys['Number'] = (T 'Tr_Number')
    }

    $keyArray = @($transformTypeKeys.Keys)
    foreach ($k in $keyArray) {
        [void]$cboType.Items.Add($transformTypeKeys[$k])
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

    $cboParam1 = New-Object Windows.Forms.ComboBox
    $cboParam1.DropDownStyle = 'DropDownList'
    $cboParam1.Location = New-Object Drawing.Point(160, 57)
    $cboParam1.Size = New-Object Drawing.Size(320, 25)
    $cboParam1.Visible = $false
    $form.Controls.Add($cboParam1)

    $cboType.Add_SelectedIndexChanged({
            foreach ($l in $paramLabels) { $l.Visible = $false }
            foreach ($c in $paramControls) { $c.Visible = $false; $c.Text = '' }
            $cboParam1.Visible = $false
            $cboParam1.Items.Clear()

            if ($cboType.SelectedIndex -lt 0) { return }
            $selectedKey = $keyArray[$cboType.SelectedIndex]

            switch ($selectedKey) {
                'Substring' {
                    $paramLabels[0].Text = (T 'Tr_PosStart'); $paramLabels[0].Visible = $true
                    $paramControls[0].Text = '0'; $paramControls[0].Visible = $true
                    $paramLabels[1].Text = (T 'Tr_CharCount'); $paramLabels[1].Visible = $true
                    $paramControls[1].Visible = $true
                }
                'DateFormat' {
                    $paramLabels[0].Text = (T 'Tr_FmtIn'); $paramLabels[0].Visible = $true
                    $paramControls[0].Text = $defaultDateInputFormat; $paramControls[0].Visible = $true
                    $paramLabels[1].Text = (T 'Tr_FmtOut'); $paramLabels[1].Visible = $true
                    $paramControls[1].Visible = $true
                }
                'Replace' {
                    $paramLabels[0].Text = (T 'Tr_Search'); $paramLabels[0].Visible = $true
                    $paramControls[0].Visible = $true
                    $paramLabels[1].Text = (T 'Tr_NewTxt'); $paramLabels[1].Visible = $true
                    $paramControls[1].Visible = $true
                }
                'Transliterate' {
                    # No params needed
                }
                'RegexReplace' {
                    $paramLabels[0].Text = (T 'Tr_RegexPattern'); $paramLabels[0].Visible = $true
                    $paramControls[0].Visible = $true
                    $paramLabels[1].Text = (T 'Tr_Replacement'); $paramLabels[1].Visible = $true
                    $paramControls[1].Visible = $true
                }
                'Expression' {
                    $paramLabels[0].Text = (T 'Tr_ExprCode'); $paramLabels[0].Visible = $true
                    $paramControls[0].Text = '$_'; $paramControls[0].Visible = $true
                }
                'Sequence' {
                    $paramLabels[0].Text = (T 'Tr_SeqStart'); $paramLabels[0].Visible = $true
                    $paramControls[0].Text = '1'; $paramControls[0].Visible = $true
                    $paramLabels[1].Text = (T 'Tr_SeqStep'); $paramLabels[1].Visible = $true
                    $paramControls[1].Text = '1'; $paramControls[1].Visible = $true
                    $paramLabels[2].Text = (T 'Tr_SeqWidth'); $paramLabels[2].Visible = $true
                    $paramControls[2].Text = '3'; $paramControls[2].Visible = $true
                }
                'Case' {
                    $paramLabels[0].Text = (T 'Tr_Mode'); $paramLabels[0].Visible = $true
                    $cboParam1.Items.AddRange(@((T 'Tr_Upper'), (T 'Tr_Lower'), (T 'Tr_Title')))
                    $cboParam1.SelectedIndex = 0
                    $cboParam1.Visible = $true
                }
                'Pad' {
                    $paramLabels[0].Text = (T 'Tr_Side'); $paramLabels[0].Visible = $true
                    $cboParam1.Items.AddRange(@((T 'Tr_Left'), (T 'Tr_Right')))
                    $cboParam1.SelectedIndex = 0
                    $cboParam1.Visible = $true
                    $paramLabels[1].Text = (T 'Tr_PadChar'); $paramLabels[1].Visible = $true
                    $paramControls[1].Text = '0'; $paramControls[1].Visible = $true
                    $paramLabels[2].Text = (T 'Tr_TargetLen'); $paramLabels[2].Visible = $true
                    $paramControls[2].Visible = $true
                }
                'Number' {
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

    $btnAdd = New-Object Windows.Forms.Button
    $btnAdd.Text = (T 'Btn_Add')
    $btnAdd.Location = New-Object Drawing.Point(400, 235)
    $btnAdd.Size = New-Object Drawing.Size(80, 30)
    $form.Controls.Add($btnAdd)

    $script:transformDialogResult = $null

    $btnAdd.Add_Click({
            try {
                if ($cboType.SelectedIndex -lt 0) { throw (T 'Err_SelTransform') }
                $selectedKey = $keyArray[$cboType.SelectedIndex]

                $result = switch ($selectedKey) {
                    'Substring' {
                        $s = [int]$paramControls[0].Text
                        $l = [int]$paramControls[1].Text
                        if ($l -le 0) { throw (T 'Err_CharCount') }
                        [pscustomobject]@{
                            Type = 'Substring'; Start = $s; Length = $l
                            Display = "$(T 'Disp_Substr') $s, $l $(T 'Disp_Chars')"
                        }
                    }
                    'DateFormat' {
                        $inf = $paramControls[0].Text.Trim()
                        $outf = $paramControls[1].Text.Trim()
                        if (-not $inf -or -not $outf) { throw (T 'Err_BothFmt') }
                        [pscustomobject]@{
                            Type = 'DateFormat'; InputFormat = $inf; OutputFormat = $outf
                            Display = "$(T 'Disp_Date') $inf → $outf"
                        }
                    }
                    'Replace' {
                        $old = $paramControls[0].Text
                        $new = $paramControls[1].Text
                        if ($old -eq '') { throw (T 'Err_SearchTxt') }
                        [pscustomobject]@{
                            Type = 'Replace'; OldText = $old; NewText = $new
                            Display = "$(T 'Disp_Replace') '$old' → '$new'"
                        }
                    }
                    'Transliterate' {
                        [pscustomobject]@{
                            Type    = 'Transliterate'
                            Display = "$(T 'Disp_Transliterate')"
                        }
                    }
                    'RegexReplace' {
                        $pat = $paramControls[0].Text
                        $rep = $paramControls[1].Text
                        if (-not $pat) { throw (T 'Err_SearchTxt') }
                        [pscustomobject]@{
                            Type = 'RegexReplace'; Pattern = $pat; Replacement = $rep
                            Display = "Regex: '$pat' → '$rep'"
                        }
                    }
                    'Expression' {
                        $expr = $paramControls[0].Text.Trim()
                        if (-not $expr) { throw (T 'Err_SearchTxt') }
                        [pscustomobject]@{
                            Type = 'Expression'; Expression = $expr
                            Display = "Expr: $expr"
                        }
                    }
                    'Sequence' {
                        $start = [int]$paramControls[0].Text
                        $step = [int]$paramControls[1].Text
                        $width = [int]$paramControls[2].Text
                        [pscustomobject]@{
                            Type = 'Sequence'; Start = $start; Step = $step; Width = $width; CurrentIndex = $start
                            Display = "Seq: Start=$start Step=$step W=$width"
                        }
                    }
                    'Case' {
                        if ($cboParam1.SelectedIndex -lt 0) { throw (T 'Err_SelMode') }
                        $mode = @('Upper', 'Lower', 'Title')[$cboParam1.SelectedIndex]
                        $modeText = [string]$cboParam1.SelectedItem
                        [pscustomobject]@{
                            Type = 'Case'; Mode = $mode
                            Display = "$(T 'Disp_Case') $modeText"
                        }
                    }
                    'Pad' {
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
                    'Number' {
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
