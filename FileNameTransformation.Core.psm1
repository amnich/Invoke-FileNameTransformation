Set-StrictMode -Version 2.0

$script:InvariantCulture = [Globalization.CultureInfo]::InvariantCulture
$script:DateFormats = @(
    'yyyyMMdd',
    'yyyy-MM-dd', 'yyyy.MM.dd', 'yyyy_MM_dd',
    'dd-MM-yyyy', 'dd.MM.yyyy', 'dd_MM_yyyy',
    'yyyy-MM', 'yyyy.MM', 'yyyy_MM',
    'MM-yyyy', 'MM.yyyy', 'MM_yyyy',
    'yyMMdd', 'ddMMyy',
    'MM-dd', 'MM.dd', 'MM_dd',
    'dd-MM', 'dd.MM', 'dd_MM',
    'yyyy-MM-ddTHH:mm:ss', 'yyyy-MM-ddTHH:mm:ssZ', 'yyyy-MM-ddTHH:mm:sszzz'
)

function New-FNTTypeCandidate {
    param(
        [Parameter(Mandatory)][string]$TypeId,
        [string]$Format,
        [string]$RuleId,
        [bool]$AllowComposite = $false
    )

    [pscustomobject]@{
        TypeId         = $TypeId
        Format         = $Format
        RuleId         = $RuleId
        AllowComposite = $AllowComposite
    }
}

function Get-FNTTypeCandidates {
    [CmdletBinding()]
    param(
        [AllowEmptyString()][string]$Value,
        [object[]]$CustomTypeRules = @()
    )

    $candidates = New-Object System.Collections.Generic.List[object]

    $integerValue = 0L
    if ($Value -match '^[+-]?\d+$' -and
        [long]::TryParse($Value, [Globalization.NumberStyles]::Integer, $script:InvariantCulture, [ref]$integerValue)) {
        $candidates.Add((New-FNTTypeCandidate -TypeId 'Integer'))
    }

    $decimalValue = 0D
    if ($Value -match '^[+-]?\d+\.\d+$' -and
        [decimal]::TryParse($Value, [Globalization.NumberStyles]::Number, $script:InvariantCulture, [ref]$decimalValue)) {
        $candidates.Add((New-FNTTypeCandidate -TypeId 'Decimal'))
    }

    foreach ($format in $script:DateFormats) {
        $dateValue = [datetime]::MinValue
        if ([datetime]::TryParseExact(
                $Value,
                $format,
                $script:InvariantCulture,
                [Globalization.DateTimeStyles]::None,
                [ref]$dateValue
            )) {
            $candidates.Add((New-FNTTypeCandidate -TypeId 'DateTime' -Format $format -AllowComposite $true))
        }
    }

    $guidValue = [guid]::Empty
    if ([guid]::TryParse($Value, [ref]$guidValue)) {
        $candidates.Add((New-FNTTypeCandidate -TypeId 'Guid' -AllowComposite $true))
    }

    $versionValue = $null
    if ($Value -match '^\d+(\.\d+){1,3}$' -and [version]::TryParse($Value, [ref]$versionValue)) {
        $candidates.Add((New-FNTTypeCandidate -TypeId 'Version' -AllowComposite $true))
    }

    foreach ($rule in @($CustomTypeRules)) {
        if ($null -eq $rule -or ($rule.PSObject.Properties['Enabled'] -and -not [bool]$rule.Enabled)) {
            continue
        }
        if (-not $rule.PSObject.Properties['Id'] -or -not $rule.PSObject.Properties['Pattern']) {
            continue
        }

        try {
            $match = [regex]::Match($Value, [string]$rule.Pattern)
        }
        catch {
            throw "Invalid custom type rule '$($rule.Id)': $($_.Exception.Message)"
        }

        if ($match.Success -and $match.Index -eq 0 -and $match.Length -eq $Value.Length) {
            $allowComposite = $rule.PSObject.Properties['AllowComposite'] -and [bool]$rule.AllowComposite
            $candidates.Add((New-FNTTypeCandidate -TypeId "Custom:$($rule.Id)" -RuleId ([string]$rule.Id) -AllowComposite $allowComposite))
        }
    }

    return $candidates.ToArray()
}

function Get-FNTLexicalTokens {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$Name,
        [Parameter(Mandatory)][string]$Pattern
    )

    try {
        $regex = [regex]::new($Pattern)
    }
    catch {
        throw "Invalid tokenizer regex: $($_.Exception.Message)"
    }

    if ($regex.GetGroupNames() -notcontains 'sep') {
        throw "Tokenizer regex must define the named group 'sep'."
    }

    $matches = @($regex.Matches($Name))
    $tokens = New-Object System.Collections.Generic.List[object]
    $expectedIndex = 0

    foreach ($match in $matches) {
        if ($match.Length -eq 0) {
            throw "Tokenizer regex produced a zero-length match at position $($match.Index)."
        }
        if ($match.Index -ne $expectedIndex) {
            throw "Tokenizer regex did not match the filename at position $expectedIndex."
        }

        $tokens.Add([pscustomobject]@{
                Value       = $match.Value
                IsSeparator = $match.Groups['sep'].Success
                Start       = $match.Index
                Length      = $match.Length
            })
        $expectedIndex = $match.Index + $match.Length
    }

    if ($expectedIndex -ne $Name.Length) {
        throw "Tokenizer regex did not match the filename at position $expectedIndex."
    }

    return $tokens.ToArray()
}

function New-FNTValueToken {
    param(
        [Parameter(Mandatory)][string]$Value,
        [Parameter(Mandatory)][int]$Start,
        [Parameter(Mandatory)][int]$Length,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Candidates,
        [Parameter(Mandatory)][object[]]$LexicalParts
    )

    $typeIds = @($Candidates | ForEach-Object { $_.TypeId } | Select-Object -Unique)
    $detectedTypeId = if ($Candidates.Count -eq 0) {
        'Text'
    }
    elseif ($Candidates.Count -eq 1) {
        [string]$Candidates[0].TypeId
    }
    else {
        'Ambiguous'
    }

    [pscustomobject]@{
        Value          = $Value
        IsSeparator    = $false
        Start          = $Start
        Length         = $Length
        DetectedTypeId = $detectedTypeId
        CandidateTypes = @($Candidates)
        CandidateIds   = @($typeIds)
        IsAmbiguous    = $Candidates.Count -gt 1
        LexicalParts   = @($LexicalParts)
    }
}

function Get-FNTTokens {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$Name,
        [string]$Pattern = '(?<value>[^_\-\s]+)|(?<sep>[_\-\s]+)',
        [object[]]$CustomTypeRules = @()
    )

    $lexicalTokens = @(Get-FNTLexicalTokens -Name $Name -Pattern $Pattern)
    $result = New-Object System.Collections.Generic.List[object]
    $index = 0

    while ($index -lt $lexicalTokens.Count) {
        $current = $lexicalTokens[$index]
        if ($current.IsSeparator) {
            $result.Add([pscustomobject]@{
                    Value          = $current.Value
                    IsSeparator    = $true
                    Start          = $current.Start
                    Length         = $current.Length
                    DetectedTypeId = 'Separator'
                    CandidateTypes = @()
                    CandidateIds   = @()
                    IsAmbiguous    = $false
                    LexicalParts   = @($current)
                })
            $index++
            continue
        }

        $selectedEnd = $index
        $selectedValue = $current.Value
        $selectedCandidates = @(Get-FNTTypeCandidates -Value $current.Value -CustomTypeRules $CustomTypeRules)

        for ($end = $lexicalTokens.Count - 1; $end -gt $index; $end--) {
            if ($lexicalTokens[$end].IsSeparator) { continue }

            $spanStart = $current.Start
            $spanLength = ($lexicalTokens[$end].Start + $lexicalTokens[$end].Length) - $spanStart
            $spanValue = $Name.Substring($spanStart, $spanLength)
            $spanCandidates = @(Get-FNTTypeCandidates -Value $spanValue -CustomTypeRules $CustomTypeRules |
                Where-Object { $_.AllowComposite })
            if ($spanCandidates.Count -gt 0) {
                $selectedEnd = $end
                $selectedValue = $spanValue
                $selectedCandidates = $spanCandidates
                break
            }
        }

        $selectedParts = @($lexicalTokens[$index..$selectedEnd])
        $length = ($selectedParts[-1].Start + $selectedParts[-1].Length) - $current.Start
        $result.Add((New-FNTValueToken -Value $selectedValue -Start $current.Start -Length $length -Candidates $selectedCandidates -LexicalParts $selectedParts))
        $index = $selectedEnd + 1
    }

    return $result.ToArray()
}

function Get-FNTFieldInference {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][object[]]$Tokens
    )

    $valueTokens = @($Tokens | Where-Object { -not $_.IsSeparator })
    if ($valueTokens.Count -eq 0) {
        throw 'Field inference requires at least one value token.'
    }

    $candidateKey = {
        param($candidate)
        "$($candidate.TypeId)|$($candidate.Format)|$($candidate.RuleId)"
    }

    $commonKeys = @{}
    foreach ($candidate in @($valueTokens[0].CandidateTypes)) {
        $commonKeys[(& $candidateKey $candidate)] = $true
    }

    for ($index = 1; $index -lt $valueTokens.Count; $index++) {
        $currentKeys = @{}
        foreach ($candidate in @($valueTokens[$index].CandidateTypes)) {
            $currentKeys[(& $candidateKey $candidate)] = $true
        }
        foreach ($key in @($commonKeys.Keys)) {
            if (-not $currentKeys.ContainsKey($key)) {
                $commonKeys.Remove($key)
            }
        }
    }

    $commonCandidates = @($valueTokens[0].CandidateTypes | Where-Object {
            $commonKeys.ContainsKey((& $candidateKey $_))
        })
    if ($commonCandidates.Count -eq 1) {
        return [pscustomobject]@{
            DetectedTypeId = [string]$commonCandidates[0].TypeId
            CandidateTypes = $commonCandidates
            IsAmbiguous    = $false
            Reason         = $null
        }
    }
    if ($commonCandidates.Count -gt 1) {
        return [pscustomobject]@{
            DetectedTypeId = 'Ambiguous'
            CandidateTypes = $commonCandidates
            IsAmbiguous    = $true
            Reason         = 'MultipleCandidates'
        }
    }

    $allCandidates = @($valueTokens | ForEach-Object { @($_.CandidateTypes) })
    if ($allCandidates.Count -eq 0) {
        return [pscustomobject]@{
            DetectedTypeId = 'Text'
            CandidateTypes = @()
            IsAmbiguous    = $false
            Reason         = $null
        }
    }

    $commonTypeIds = @($valueTokens[0].CandidateIds)
    for ($index = 1; $index -lt $valueTokens.Count; $index++) {
        $currentTypeIds = @($valueTokens[$index].CandidateIds)
        $commonTypeIds = @($commonTypeIds | Where-Object { $currentTypeIds -contains $_ })
    }
    $commonTypeIds = @($commonTypeIds | Select-Object -Unique)

    $detectedTypeId = if ($commonTypeIds.Count -eq 1) { [string]$commonTypeIds[0] } else { 'Ambiguous' }
    $relevantCandidates = if ($commonTypeIds.Count -eq 1) {
        @($allCandidates | Where-Object { $_.TypeId -eq $commonTypeIds[0] } |
            Sort-Object TypeId, Format, RuleId -Unique)
    }
    else {
        @($allCandidates | Sort-Object TypeId, Format, RuleId -Unique)
    }

    return [pscustomobject]@{
        DetectedTypeId = $detectedTypeId
        CandidateTypes = $relevantCandidates
        IsAmbiguous    = $true
        Reason         = 'InconsistentCandidates'
    }
}

function Test-FNTValueType {
    [CmdletBinding()]
    param(
        [AllowEmptyString()][string]$Value,
        [Parameter(Mandatory)][string]$TypeId,
        [string]$Format,
        [object[]]$CustomTypeRules = @()
    )

    if ($TypeId -in @('Auto', 'Text')) { return $true }
    $candidates = @(Get-FNTTypeCandidates -Value $Value -CustomTypeRules $CustomTypeRules)
    foreach ($candidate in $candidates) {
        if ($candidate.TypeId -ne $TypeId) { continue }
        if ($TypeId -ne 'DateTime' -or [string]::IsNullOrWhiteSpace($Format) -or $candidate.Format -eq $Format) {
            return $true
        }
    }
    return $false
}

function Match-FNTNamePattern {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$Name,
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][object[]]$PatternTokens,
        [Parameter(Mandatory)][string]$TokenizerPattern,
        [hashtable]$FieldTypes = @{},
        [object[]]$CustomTypeRules = @()
    )

    $actualTokens = @(Get-FNTTokens -Name $Name -Pattern $TokenizerPattern -CustomTypeRules $CustomTypeRules)
    if ($actualTokens.Count -ne $PatternTokens.Count) {
        throw "Pattern mismatch for '$Name': expected $($PatternTokens.Count) tokens, found $($actualTokens.Count)."
    }

    $values = @{}
    for ($index = 0; $index -lt $PatternTokens.Count; $index++) {
        $expected = $PatternTokens[$index]
        $actual = $actualTokens[$index]
        if ([bool]$expected.IsSeparator -ne [bool]$actual.IsSeparator) {
            $expectedKind = if ($expected.IsSeparator) { 'separator' } else { 'value' }
            $actualKind = if ($actual.IsSeparator) { 'separator' } else { 'value' }
            throw "Pattern mismatch for '$Name' at token $($index + 1): expected $expectedKind, found $actualKind '$($actual.Value)'."
        }
        if ($expected.IsSeparator) {
            if ($expected.Value -cne $actual.Value) {
                throw "Pattern mismatch for '$Name' at token $($index + 1): expected separator '$($expected.Value)', found '$($actual.Value)'."
            }
            continue
        }

        if ($FieldTypes.ContainsKey($index)) {
            $fieldType = $FieldTypes[$index]
            if (-not (Test-FNTValueType -Value $actual.Value -TypeId ([string]$fieldType.TypeId) -Format ([string]$fieldType.Format) -CustomTypeRules $CustomTypeRules)) {
                $formatText = if ($fieldType.Format) { " format '$($fieldType.Format)'" } else { '' }
                throw "Pattern mismatch for '$Name' at token $($index + 1): value '$($actual.Value)' is not type '$($fieldType.TypeId)'$formatText."
            }
        }
        $values[$index] = $actual.Value
    }

    [pscustomobject]@{
        Values = $values
        Tokens = $actualTokens
    }
}

Export-ModuleMember -Function Get-FNTLexicalTokens, Get-FNTTypeCandidates, Get-FNTTokens, Get-FNTFieldInference, Test-FNTValueType, Match-FNTNamePattern