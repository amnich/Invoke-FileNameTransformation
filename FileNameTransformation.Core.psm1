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

function New-FNTException {
    param(
        [Parameter(Mandatory)][string]$Code,
        [Parameter(Mandatory)][string]$Message,
        [hashtable]$Details = @{}
    )

    $exception = [InvalidOperationException]::new($Message)
    $exception.Data['FNTCode'] = $Code
    foreach ($key in $Details.Keys) {
        $exception.Data[$key] = $Details[$key]
    }
    return $exception
}

function New-FNTTypeCandidate {
    param(
        [Parameter(Mandatory)][string]$TypeId,
        [string]$Format,
        [string]$RuleId,
        [string]$DisplayName,
        [bool]$AllowComposite = $false
    )

    [pscustomobject]@{
        TypeId         = $TypeId
        Format         = $Format
        RuleId         = $RuleId
        DisplayName    = $DisplayName
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
            $displayName = if ($rule.PSObject.Properties['DisplayName'] -and $rule.DisplayName) {
                [string]$rule.DisplayName
            }
            else {
                [string]$rule.Id
            }
            $candidates.Add((New-FNTTypeCandidate -TypeId "Custom:$($rule.Id)" -RuleId ([string]$rule.Id) -DisplayName $displayName -AllowComposite $allowComposite))
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
        throw (New-FNTException -Code 'Tokenizer.InvalidRegex' -Message "Invalid tokenizer regex: $($_.Exception.Message)" -Details @{
                Reason = $_.Exception.Message
            })
    }

    if ($regex.GetGroupNames() -notcontains 'sep') {
        throw (New-FNTException -Code 'Tokenizer.MissingSeparatorGroup' -Message "Tokenizer regex must define the named group 'sep'.")
    }

    $matches = @($regex.Matches($Name))
    $tokens = New-Object System.Collections.Generic.List[object]
    $expectedIndex = 0

    foreach ($match in $matches) {
        if ($match.Length -eq 0) {
            throw (New-FNTException -Code 'Tokenizer.ZeroLength' -Message "Tokenizer regex produced a zero-length match at position $($match.Index)." -Details @{
                    Position = $match.Index
                })
        }
        if ($match.Index -ne $expectedIndex) {
            throw (New-FNTException -Code 'Tokenizer.IncompleteCoverage' -Message "Tokenizer regex did not match the filename at position $expectedIndex." -Details @{
                    Position = $expectedIndex
                })
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
        throw (New-FNTException -Code 'Tokenizer.IncompleteCoverage' -Message "Tokenizer regex did not match the filename at position $expectedIndex." -Details @{
            Position = $expectedIndex
            })
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

function Get-FNTTokenSignature {
    [CmdletBinding()]
    param([Parameter(Mandatory)][ValidateNotNullOrEmpty()][object[]]$Tokens)

    return (($Tokens | ForEach-Object {
                if ($_.IsSeparator) { 'S:' + $_.Value + '|' }
                else { 'T:' + $_.DetectedTypeId + '|' }
            }) -join '')
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
        throw (New-FNTException -Code 'Pattern.TokenCount' -Message "Pattern mismatch for '$Name': expected $($PatternTokens.Count) tokens, found $($actualTokens.Count)." -Details @{
            Name = $Name; Expected = $PatternTokens.Count; Actual = $actualTokens.Count
            })
    }

    $values = @{}
    for ($index = 0; $index -lt $PatternTokens.Count; $index++) {
        $expected = $PatternTokens[$index]
        $actual = $actualTokens[$index]
        if ([bool]$expected.IsSeparator -ne [bool]$actual.IsSeparator) {
            $expectedKind = if ($expected.IsSeparator) { 'separator' } else { 'value' }
            $actualKind = if ($actual.IsSeparator) { 'separator' } else { 'value' }
            throw (New-FNTException -Code 'Pattern.TokenKind' -Message "Pattern mismatch for '$Name' at token $($index + 1), offset $($actual.Start): expected $expectedKind, found $actualKind '$($actual.Value)'." -Details @{
                    Name = $Name; Token = $index + 1; Offset = $actual.Start; Expected = $expectedKind; Actual = $actualKind; Value = $actual.Value
                })
        }
        if ($expected.IsSeparator) {
            if ($expected.Value -cne $actual.Value) {
                throw (New-FNTException -Code 'Pattern.Separator' -Message "Pattern mismatch for '$Name' at token $($index + 1), offset $($actual.Start): expected separator '$($expected.Value)', found '$($actual.Value)'." -Details @{
                        Name = $Name; Token = $index + 1; Offset = $actual.Start; Expected = $expected.Value; Actual = $actual.Value
                    })
            }
            continue
        }

        if ($FieldTypes.ContainsKey($index)) {
            $fieldType = $FieldTypes[$index]
            if (-not (Test-FNTValueType -Value $actual.Value -TypeId ([string]$fieldType.TypeId) -Format ([string]$fieldType.Format) -CustomTypeRules $CustomTypeRules)) {
                $formatText = if ($fieldType.Format) { " format '$($fieldType.Format)'" } else { '' }
                throw (New-FNTException -Code 'Pattern.Type' -Message "Pattern mismatch for '$Name' at token $($index + 1), offset $($actual.Start): value '$($actual.Value)' is not type '$($fieldType.TypeId)'$formatText." -Details @{
                        Name = $Name; Token = $index + 1; Offset = $actual.Start; Value = $actual.Value; TypeId = $fieldType.TypeId; Format = $fieldType.Format
                    })
            }
        }
        $values[$index] = $actual.Value
    }

    [pscustomobject]@{
        Values = $values
        Tokens = $actualTokens
    }
}

function Copy-FNTObject {
    param([AllowNull()]$InputObject)

    if ($null -eq $InputObject) { return $null }
    return ($InputObject | ConvertTo-Json -Depth 20 | ConvertFrom-Json)
}

function ConvertFrom-FNTLegacyTypeLabel {
    [CmdletBinding()]
    param([AllowNull()][string]$TypeLabel)

    if ([string]::IsNullOrWhiteSpace($TypeLabel)) { return 'Text' }
    if ($TypeLabel -match 'Ambig|Niejedno|Mehrdeu') { return 'Ambiguous' }
    if ($TypeLabel -match 'GUID') { return 'Guid' }
    if ($TypeLabel -match 'Version|Wersja') { return 'Version' }
    if ($TypeLabel -match 'Decimal|dzies|Dezimal') { return 'Decimal' }
    if ($TypeLabel -match 'Date|Data|Datum') { return 'DateTime' }
    if ($TypeLabel -match 'Number|Liczba|Zahl') { return 'Integer' }
    return 'Text'
}

function ConvertTo-FNTConfig {
    [CmdletBinding()]
    param([AllowNull()]$Config)

    $normalized = Copy-FNTObject $Config
    if ($null -eq $normalized) {
        $normalized = [pscustomobject][ordered]@{}
    }
    if (-not $normalized.PSObject.Properties['Version']) {
        $normalized | Add-Member -NotePropertyName Version -NotePropertyValue 2
    }
    else {
        $normalized.Version = 2
    }
    if (-not $normalized.PSObject.Properties['Language'] -or $normalized.Language -notin @('PL', 'EN', 'DE')) {
        if ($normalized.PSObject.Properties['Language']) { $normalized.Language = 'PL' }
        else { $normalized | Add-Member -NotePropertyName Language -NotePropertyValue 'PL' }
    }
    if (-not $normalized.PSObject.Properties['CustomTypeRules']) {
        $normalized | Add-Member -NotePropertyName CustomTypeRules -NotePropertyValue @()
    }
    else {
        $normalized.CustomTypeRules = @($normalized.CustomTypeRules)
    }
    return $normalized
}

function Set-FNTConfigLanguage {
    [CmdletBinding()]
    param(
        [AllowNull()]$Config,
        [Parameter(Mandatory)][ValidateSet('PL', 'EN', 'DE')][string]$Language
    )

    $normalized = ConvertTo-FNTConfig $Config
    $normalized.Language = $Language
    return $normalized
}

function Test-FNTCustomTypeRules {
    [CmdletBinding()]
    param([object[]]$Rules = @())

    $validRules = New-Object System.Collections.Generic.List[object]
    $errors = New-Object System.Collections.Generic.List[object]
    $seenIds = @{}

    for ($index = 0; $index -lt @($Rules).Count; $index++) {
        $rule = @($Rules)[$index]
        $ruleNumber = $index + 1
        if ($null -eq $rule) {
            $errors.Add([pscustomobject]@{ Index = $index; Id = $null; Message = "Rule $ruleNumber is null." })
            continue
        }

        $id = if ($rule.PSObject.Properties['Id']) { ([string]$rule.Id).Trim() } else { '' }
        if ($id -notmatch '^[A-Za-z][A-Za-z0-9_.-]*$') {
            $errors.Add([pscustomobject]@{ Index = $index; Id = $id; Message = "Rule $ruleNumber has an invalid ID '$id'." })
            continue
        }
        $idKey = $id.ToUpperInvariant()
        if ($seenIds.ContainsKey($idKey)) {
            $errors.Add([pscustomobject]@{ Index = $index; Id = $id; Message = "Rule $ruleNumber duplicates ID '$id'." })
            continue
        }
        $seenIds[$idKey] = $true

        $pattern = if ($rule.PSObject.Properties['Pattern']) { [string]$rule.Pattern } else { '' }
        if ([string]::IsNullOrWhiteSpace($pattern)) {
            $errors.Add([pscustomobject]@{ Index = $index; Id = $id; Message = "Rule '$id' has an empty pattern." })
            continue
        }
        try {
            [void][regex]::new($pattern)
        }
        catch {
            $errors.Add([pscustomobject]@{ Index = $index; Id = $id; Message = "Rule '$id' has an invalid regex: $($_.Exception.Message)" })
            continue
        }

        $displayName = if ($rule.PSObject.Properties['DisplayName'] -and
            -not [string]::IsNullOrWhiteSpace([string]$rule.DisplayName)) {
            ([string]$rule.DisplayName).Trim()
        }
        else {
            $id
        }
        $enabled = if ($rule.PSObject.Properties['Enabled']) { [bool]$rule.Enabled } else { $true }
        $allowComposite = $rule.PSObject.Properties['AllowComposite'] -and [bool]$rule.AllowComposite

        $validRules.Add([pscustomobject][ordered]@{
                Id             = $id
                DisplayName    = $displayName
                Pattern        = $pattern
                Enabled        = $enabled
                AllowComposite = $allowComposite
            })
    }

    return [pscustomobject]@{
        ValidRules = $validRules.ToArray()
        Errors     = $errors.ToArray()
    }
}

function ConvertTo-FNTProfile {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Profile)

    $normalized = Copy-FNTObject $Profile
    if (-not $normalized.PSObject.Properties['SchemaVersion']) {
        $normalized | Add-Member -NotePropertyName SchemaVersion -NotePropertyValue 2
    }
    else {
        $normalized.SchemaVersion = 2
    }
    if (-not $normalized.PSObject.Properties['TokenRegex'] -or
        [string]::IsNullOrWhiteSpace([string]$normalized.TokenRegex)) {
        $defaultPattern = '(?<value>[^_\-\s]+)|(?<sep>[_\-\s]+)'
        if ($normalized.PSObject.Properties['TokenRegex']) { $normalized.TokenRegex = $defaultPattern }
        else { $normalized | Add-Member -NotePropertyName TokenRegex -NotePropertyValue $defaultPattern }
    }

    foreach ($propertyName in @('Fields', 'Mappings', 'OutputParts')) {
        if (-not $normalized.PSObject.Properties[$propertyName]) {
            $normalized | Add-Member -NotePropertyName $propertyName -NotePropertyValue @()
        }
        else {
            $normalized.$propertyName = @($normalized.$propertyName)
        }
    }
    if (-not $normalized.PSObject.Properties['KeepExtension']) {
        $normalized | Add-Member -NotePropertyName KeepExtension -NotePropertyValue $true
    }
    if (-not $normalized.PSObject.Properties['NewExtension']) {
        $normalized | Add-Member -NotePropertyName NewExtension -NotePropertyValue ''
    }

    foreach ($field in @($normalized.Fields)) {
        if (-not $field.PSObject.Properties['Transforms']) {
            $field | Add-Member -NotePropertyName Transforms -NotePropertyValue @()
        }
        else {
            $field.Transforms = @($field.Transforms)
        }
        if (-not $field.PSObject.Properties['IsVirtual']) {
            $field | Add-Member -NotePropertyName IsVirtual -NotePropertyValue $false
        }
        if (-not $field.PSObject.Properties['PartIndex']) {
            $partIndex = if ($field.PSObject.Properties['Index']) { [int]$field.Index } else { -1 }
            $field | Add-Member -NotePropertyName PartIndex -NotePropertyValue $partIndex
        }
        if (-not $field.PSObject.Properties['DisplayIndex']) {
            $displayIndex = if ($field.IsVirtual) { 'V' } else { [string]$field.PartIndex }
            $field | Add-Member -NotePropertyName DisplayIndex -NotePropertyValue $displayIndex
        }
        if (-not $field.PSObject.Properties['DetectedTypeId']) {
            $typeLabel = if ($field.PSObject.Properties['DetectedType']) { [string]$field.DetectedType } else { '' }
            $field | Add-Member -NotePropertyName DetectedTypeId -NotePropertyValue (ConvertFrom-FNTLegacyTypeLabel $typeLabel)
        }
        if (-not $field.PSObject.Properties['CandidateTypes']) {
            $field | Add-Member -NotePropertyName CandidateTypes -NotePropertyValue @()
        }
        else {
            $field.CandidateTypes = @($field.CandidateTypes)
        }
        if (-not $field.PSObject.Properties['IsAmbiguous']) {
            $field | Add-Member -NotePropertyName IsAmbiguous -NotePropertyValue ($field.DetectedTypeId -eq 'Ambiguous')
        }
        if (-not $field.PSObject.Properties['SelectedTypeId']) {
            $field | Add-Member -NotePropertyName SelectedTypeId -NotePropertyValue 'Auto'
        }
        if (-not $field.PSObject.Properties['SelectedFormat']) {
            $field | Add-Member -NotePropertyName SelectedFormat -NotePropertyValue $null
        }
    }

    return $normalized
}

function Get-FNTFileMetadata {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path
    )

    $result = [ordered]@{
        CreationDate    = $null
        CreationDateStr = ''
        Author          = ''
        AuthorSurname7  = ''
        AuthorInitial   = ''
        AuthorSegment   = ''
        Title           = ''
        LastModified    = $null
        LastModifiedStr = ''
        DateTaken       = ''
        DateTakenStr    = ''
        Dimensions      = ''
        Camera          = ''
        AudioArtist     = ''
        AudioTitle      = ''
        AudioAlbum      = ''
        AudioYear       = ''
        DocCreator      = ''
        DocTitle        = ''
        DocSubject      = ''
        HashMD5         = ''
        HashSHA256      = ''
    }

    if (Test-Path -LiteralPath $Path -PathType Leaf) {
        try {
            $fi = Get-Item -LiteralPath $Path -ErrorAction Stop
            $result.CreationDate = $fi.CreationTime
            $result.CreationDateStr = $fi.CreationTime.ToString('yyyyMMdd')
            $result.LastModified = $fi.LastWriteTime
            $result.LastModifiedStr = $fi.LastWriteTime.ToString('yyyyMMdd')

            # Calculate content hashes
            try {
                $md5 = [System.Security.Cryptography.MD5]::Create()
                $sha256 = [System.Security.Cryptography.SHA256]::Create()
                $stream = [System.IO.File]::OpenRead($Path)
                try {
                    $md5Bytes = $md5.ComputeHash($stream)
                    $result.HashMD5 = [BitConverter]::ToString($md5Bytes) -replace '-'
                    [void]$stream.Seek(0, [System.IO.SeekOrigin]::Begin)
                    $shaBytes = $sha256.ComputeHash($stream)
                    $result.HashSHA256 = [BitConverter]::ToString($shaBytes) -replace '-'
                }
                finally {
                    $stream.Dispose()
                    $md5.Dispose()
                    $sha256.Dispose()
                }
            }
            catch {}
        }
        catch {}

        # Shell COM Extended Properties (Author, Title, Audio tags)
        try {
            $shell = New-Object -ComObject Shell.Application
            $parent = Split-Path $Path -Parent
            $leaf   = Split-Path $Path -Leaf
            $folder = $shell.NameSpace($parent)
            if ($folder) {
                $item = $folder.ParseName($leaf)
                if ($item) {
                    $author = $folder.GetDetailsOf($item, 20)
                    $title  = $folder.GetDetailsOf($item, 21)
                    if ($author) { $result.Author = [string]$author.Trim() }
                    if ($title)  { $result.Title  = [string]$title.Trim() }

                    # Audio extended tags
                    $artist = $folder.GetDetailsOf($item, 13)
                    $album  = $folder.GetDetailsOf($item, 14)
                    $year   = $folder.GetDetailsOf($item, 15)
                    if (-not $year) { $year = $folder.GetDetailsOf($item, 28) }
                    if ($artist) { $result.AudioArtist = [string]$artist.Trim() }
                    if ($title)  { $result.AudioTitle  = [string]$title.Trim() }
                    if ($album)  { $result.AudioAlbum  = [string]$album.Trim() }
                    if ($year)   { $result.AudioYear   = [string]$year.Trim() }
                }
            }
        }
        catch {}

        # Image EXIF & Dimension Metadata (.jpg, .jpeg, .png, .tiff)
        $ext = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()
        if ($ext -in @('.jpg', '.jpeg', '.png', '.tif', '.tiff')) {
            try {
                Add-Type -AssemblyName PresentationCore -ErrorAction SilentlyContinue
                $fs = [System.IO.File]::OpenRead($Path)
                try {
                    $decoder = [System.Windows.Media.Imaging.BitmapDecoder]::Create($fs, [System.Windows.Media.Imaging.BitmapCreateOptions]::None, [System.Windows.Media.Imaging.BitmapCacheOption]::None)
                    if ($decoder.Frames.Count -gt 0) {
                        $frame = $decoder.Frames[0]
                        $result.Dimensions = "$($frame.PixelWidth)x$($frame.PixelHeight)"
                        $metadata = $frame.Metadata -as [System.Windows.Media.Imaging.BitmapMetadata]
                        if ($null -ne $metadata) {
                            if ($metadata.DateTaken) {
                                $dt = [datetime]::Parse($metadata.DateTaken)
                                $result.DateTaken = $dt.ToString('yyyy-MM-dd HH:mm:ss')
                                $result.DateTakenStr = $dt.ToString('yyyyMMdd')
                            }
                            if ($metadata.CameraModel) { $result.Camera = [string]$metadata.CameraModel }
                        }
                    }
                }
                finally {
                    $fs.Dispose()
                }
            }
            catch {}
        }

        # Office OpenXML Metadata (.docx, .xlsx, .pptx)
        if ($ext -in @('.docx', '.xlsx', '.pptx')) {
            try {
                Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
                $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
                try {
                    $coreEntry = $zip.Entries | Where-Object { $_.FullName -eq 'docProps/core.xml' }
                    if ($coreEntry) {
                        $stream = $coreEntry.Open()
                        $reader = [System.IO.StreamReader]::new($stream)
                        $xmlContent = $reader.ReadToEnd()
                        $reader.Dispose()
                        $stream.Dispose()
                        [xml]$xml = $xmlContent
                        if ($xml.coreProperties) {
                            if ($xml.coreProperties.creator) { $result.DocCreator = [string]$xml.coreProperties.creator }
                            if ($xml.coreProperties.title) { $result.DocTitle = [string]$xml.coreProperties.title }
                            if ($xml.coreProperties.subject) { $result.DocSubject = [string]$xml.coreProperties.subject }
                        }
                    }
                }
                finally {
                    $zip.Dispose()
                }
            }
            catch {}
        }

        if ($result.Author) {
            $authorClean = $result.Author -replace '[^\p{L}\s,-]', ''
            $parts = @($authorClean -split '\s+|,') | Where-Object { $_.Trim() }
            if ($parts.Count -ge 2) {
                if ($result.Author -match ',') {
                    $surname = $parts[0].Trim()
                    $given   = $parts[1].Trim()
                } else {
                    $given   = $parts[0].Trim()
                    $surname = $parts[-1].Trim()
                }
                if ($surname -and $given) {
                    $s7 = $surname.Substring(0, [Math]::Min(7, $surname.Length))
                    $g1 = $given.Substring(0, 1).ToUpper()
                    # Format as Titlecase surname (up to 7 chars) + 1 uppercase initial
                    $s7Formatted = (Get-Culture).TextInfo.ToTitleCase($s7.ToLower())
                    $result.AuthorSurname7 = $s7Formatted
                    $result.AuthorInitial  = $g1
                    # Combine surname + initial, then pad with '-' to length 8 if needed
                    $baseSeg = "$s7Formatted$g1"
                    if ($baseSeg.Length -lt 8) {
                        $baseSeg = $baseSeg.PadRight(8, '-')
                    }
                    $result.AuthorSegment = $baseSeg
                }
            }
            elseif ($parts.Count -eq 1 -and $parts[0].Length -ge 2) {
                $s7 = $parts[0].Substring(0, [Math]::Min(7, $parts[0].Length))
                $s7Formatted = (Get-Culture).TextInfo.ToTitleCase($s7.ToLower())
                $result.AuthorSurname7 = $s7Formatted
                $result.AuthorSegment  = $s7Formatted.PadRight(8, '-')
            }
        }
    }

    return [pscustomobject]$result
}

function Get-FNTNormalizedAuthorSegment {
    [CmdletBinding()]
    param(
        [AllowEmptyString()][string]$Raw
    )

    if ([string]::IsNullOrWhiteSpace($Raw)) {
        return 'UnknownA'
    }

    $cleanNoHyphens = $Raw -replace '[^\p{L}]', ''
    if ($cleanNoHyphens.Length -lt 2) {
        return $Raw.PadRight(8, '-')
    }

    $lastChar = $cleanNoHyphens.Substring($cleanNoHyphens.Length - 1, 1)
    $firstChar = $cleanNoHyphens.Substring(0, 1)

    $isLastUpper = ($lastChar -cmatch '[\p{Lu}]')
    $isFirstUpper = ($firstChar -cmatch '[\p{Lu}]')

    if ($isLastUpper -and ($cleanNoHyphens.Length -eq 2 -or $cleanNoHyphens.Substring(0, $cleanNoHyphens.Length - 1) -cmatch '[\p{Ll}]')) {
        # SurnameFirst: "Mnich" + "A"
        $surname = $cleanNoHyphens.Substring(0, $cleanNoHyphens.Length - 1)
        $initial = $lastChar
    }
    elseif ($isFirstUpper -and ($cleanNoHyphens.Substring(1) -cmatch '[\p{Ll}]')) {
        # GivenFirst: "A" + "Mnich" -> surname "Mnich", initial "A"
        $initial = $firstChar
        $surname = $cleanNoHyphens.Substring(1)
    }
    else {
        # Default fallback
        $surname = $cleanNoHyphens.Substring(0, [Math]::Min(7, $cleanNoHyphens.Length - 1))
        $initial = $cleanNoHyphens.Substring($cleanNoHyphens.Length - 1, 1).ToUpper()
    }

    $s7 = $surname.Substring(0, [Math]::Min(7, $surname.Length))
    $s7Formatted = (Get-Culture).TextInfo.ToTitleCase($s7.ToLower())
    $i1Formatted = $initial.ToUpper()

    $result = "$s7Formatted$i1Formatted"
    return $result.PadRight(8, '-')
}

function Test-FNTNamingConvention {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$BaseName
    )

    $violations = New-Object System.Collections.Generic.List[string]
    $segments = [ordered]@{
        Date     = ''
        Author   = ''
        DocType  = ''
        FreeText = ''
        Version  = ''
    }
    $authorAnalysis = [ordered]@{
        Raw               = ''
        DetectedFormat    = 'Unknown' # SurnameFirst, GivenFirst, TooShort, TooLong, Unknown
        SuggestedAuthor   = ''
    }

    $parts = $BaseName -split '_'

    if ($parts.Count -lt 4) {
        $violations.Add('Compliance_StructureBad')
        if ($parts.Count -eq 3) {
            $isPart0Date = ($parts[0] -match '^\d{8}$')
            $isPart1Author = ($parts[1] -match '^[\p{L}-]+$')

            if ($isPart0Date -and $isPart1Author) {
                $segments.Date     = $parts[0]
                $segments.Author   = $parts[1]
                $segments.DocType  = ''
                $segments.FreeText = $parts[2]
            }
            elseif ($isPart0Date) {
                $segments.Date     = $parts[0]
                $segments.Author   = ''
                $segments.DocType  = $parts[1]
                $segments.FreeText = $parts[2]
            }
            else {
                $segments.Author   = $parts[0]
                $segments.DocType  = $parts[1]
                $segments.FreeText = $parts[2]
            }
        }
        elseif ($parts.Count -eq 2) {
            if ($parts[0] -match '^\d{8}$') {
                $segments.Date     = $parts[0]
                $segments.FreeText = $parts[1]
            }
            elseif ($parts[0] -match '^[\p{L}-]+$') {
                $segments.Author   = $parts[0]
                $segments.FreeText = $parts[1]
            }
            else {
                $segments.FreeText = ($parts -join '_')
            }
        }
        elseif ($parts.Count -eq 1) {
            $segments.FreeText = $parts[0]
        }
    }
    else {
        # Extract version if 5th or later part matches ^v\d+$
        $hasVersion = $false
        if ($parts.Count -ge 5 -and $parts[-1] -match '^v\d+$') {
            $segments.Version = $parts[-1]
            $hasVersion = $true
            $freeTextParts = $parts[3..($parts.Count - 2)]
        }
        else {
            $freeTextParts = $parts[3..($parts.Count - 1)]
        }

        $segments.Date     = $parts[0]
        $segments.Author   = $parts[1]
        $segments.DocType  = $parts[2]
        $segments.FreeText = ($freeTextParts -join '_')
    }

    # 1. Validate Date (JJJJMMTT)
    $parsedDate = [datetime]::MinValue
    if (-not ($segments.Date -match '^\d{8}$' -and
        [datetime]::TryParseExact($segments.Date, 'yyyyMMdd', $script:InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$parsedDate))) {
        $violations.Add('Compliance_DateBad')
    }

    # 2. Validate Author (XxxxxxxY: 7 chars surname + 1 char initial = 8 chars, padded on right with '-' if shorter)
    $authorRaw = $segments.Author
    $authorAnalysis.Raw = $authorRaw

    $suggested = Get-FNTNormalizedAuthorSegment $authorRaw
    $authorAnalysis.SuggestedAuthor = $suggested

    if ($authorRaw -eq $suggested) {
        $authorAnalysis.DetectedFormat = 'SurnameFirst'
    }
    else {
        if ($authorRaw.Length -lt 8) {
            $authorAnalysis.DetectedFormat = 'TooShort'
            $violations.Add('Compliance_AuthorShort')
        }
        elseif ($authorRaw.Length -gt 8) {
            $authorAnalysis.DetectedFormat = 'TooLong'
            $violations.Add('Compliance_AuthorLong')
        }
        else {
            $authorAnalysis.DetectedFormat = 'GivenFirst'
            $violations.Add('Compliance_AuthorReversed')
        }
    }

    # 3. DocType: non-empty check
    if ([string]::IsNullOrWhiteSpace($segments.DocType)) {
        $violations.Add('Compliance_StructureBad')
    }

    # 4. FreeText: non-empty check
    if ([string]::IsNullOrWhiteSpace($segments.FreeText)) {
        $violations.Add('Compliance_StructureBad')
    }

    # 5. Version: optional, but if present must be valid v\d+
    if ($parts.Count -ge 5 -and -not $hasVersion -and $parts[-1] -like 'v*') {
        $violations.Add('Compliance_NoVersion')
    }

    return [pscustomobject]@{
        IsCompliant    = ($violations.Count -eq 0)
        Segments       = [pscustomobject]$segments
        Violations     = @($violations)
        AuthorAnalysis = [pscustomobject]$authorAnalysis
    }
}

Export-ModuleMember -Function Get-FNTLexicalTokens, Get-FNTTypeCandidates, Get-FNTTokens, Get-FNTTokenSignature, Get-FNTFieldInference, Test-FNTValueType, Match-FNTNamePattern, ConvertFrom-FNTLegacyTypeLabel, ConvertTo-FNTConfig, Set-FNTConfigLanguage, Test-FNTCustomTypeRules, ConvertTo-FNTProfile, Get-FNTFileMetadata, Test-FNTNamingConvention, Get-FNTNormalizedAuthorSegment