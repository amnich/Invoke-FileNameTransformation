Set-StrictMode -Version 2.0

$script:InvariantCulture = [Globalization.CultureInfo]::InvariantCulture
$script:ShellMetadataHeaderIndexes = @{}
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

<#
.SYNOPSIS
    Creates a structured Exception object with FNT error codes and diagnostic details.

.DESCRIPTION
    Instantiates an InvalidOperationException with attached FNT error code metadata
    and contextual key-value pairs stored in the Exception.Data collection.

.PARAMETER Code
    Unique string identifier for the error category (e.g., 'Tokenizer.InvalidRegex').

.PARAMETER Message
    Human-readable description of the error condition.

.PARAMETER Details
    Hashtable of key-value diagnostic metadata attached to the exception.

.OUTPUTS
    [System.InvalidOperationException] Structured exception object with Data entries.
#>
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

<#
.SYNOPSIS
    Constructs a data type candidate object for token classification.

.DESCRIPTION
    Returns a structured object representing a potential data type interpretation for a value token,
    including type ID, date format string, custom rule ID, display name, and composite span support.

.PARAMETER TypeId
    Identifies the semantic data type (e.g., 'Integer', 'Decimal', 'DateTime', 'Guid', 'Version', 'Custom:RuleId').

.PARAMETER Format
    Exact date format string if TypeId is 'DateTime'.

.PARAMETER RuleId
    The ID of the custom regex type rule if TypeId starts with 'Custom:'.

.PARAMETER DisplayName
    User-friendly display name for custom data types.

.PARAMETER AllowComposite
    Boolean indicating whether this candidate type can span across lexical separators.

.OUTPUTS
    [PSCustomObject] Type candidate definition.
#>
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

<#
.SYNOPSIS
    Evaluates a string value against built-in and custom data type recognizers.

.DESCRIPTION
    Analyzes an input string and detects matching data types (Integer, Decimal, DateTime in exact formats,
    GUID, Version, and Custom Regex types configured in CustomTypeRules). Returns an array of type candidates.

.PARAMETER Value
    The string token to test for type candidates.

.PARAMETER CustomTypeRules
    Array of custom type rule objects loaded from configuration.

.EXAMPLE
    Get-FNTTypeCandidates -Value '2026-01-16'
    Returns DateTime candidate with format 'yyyy-MM-dd' and AllowComposite = $true.

.OUTPUTS
    [PSCustomObject[]] Collection of matching type candidate objects.
#>
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

<#
.SYNOPSIS
    Splits a filename string into literal separator tokens and value tokens using a tokenizer regex.

.DESCRIPTION
    Executes a regex pattern containing the named capture group (?<sep>...) against a filename string.
    Ensures complete string coverage and validates that no zero-length matches occur.

.PARAMETER Name
    The raw filename string (without directory path or extension).

.PARAMETER Pattern
    Tokenizer regex pattern with named group 'sep' for separators.

.OUTPUTS
    [PSCustomObject[]] Array of lexical token objects containing Value, IsSeparator, Start, and Length.
#>
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

<#
.SYNOPSIS
    Constructs a composite value token object with detected type classification.

.DESCRIPTION
    Aggregates candidate types for a token or composite span of lexical parts, determining whether the
    token is unambiguous ('Text', 'Integer', 'DateTime', etc.) or 'Ambiguous'.

.PARAMETER Value
    The string value of the token or merged composite span.

.PARAMETER Start
    Start character index within the source filename.

.PARAMETER Length
    Character length of the token.

.PARAMETER Candidates
    Array of type candidate objects matched for this value.

.PARAMETER LexicalParts
    Array of individual lexical tokens combined to form this value.

.OUTPUTS
    [PSCustomObject] Detailed value token object.
#>
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

<#
.SYNOPSIS
    Performs full two-pass tokenization of a filename into value and separator tokens.

.DESCRIPTION
    Performs lexical tokenization followed by a second semantic pass that recognizes composite values
    (such as dates '2026-01-16' or GUIDs) spanning across punctuation separators.

.PARAMETER Name
    The filename string to tokenize.

.PARAMETER Pattern
    Tokenizer regex pattern with named group 'sep'. Defaults to splitting on underscores, hyphens, and spaces.

.PARAMETER CustomTypeRules
    Optional array of active custom type regex rules.

.OUTPUTS
    [PSCustomObject[]] Ordered collection of token objects (separators and value tokens).
#>
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

<#
.SYNOPSIS
    Generates a structural signature string representing token order and separator literals.

.DESCRIPTION
    Creates a unique string signature combining separator literal values and value token type classifications.
    Used for grouping files with identical structural patterns during analysis.

.PARAMETER Tokens
    Array of token objects returned by Get-FNTTokens.

.OUTPUTS
    [String] Token signature string (e.g. 'T:Text|S:_|T:DateTime|S:_|T:Text|').
#>
function Get-FNTTokenSignature {
    [CmdletBinding()]
    param([Parameter(Mandatory)][ValidateNotNullOrEmpty()][object[]]$Tokens)

    return (($Tokens | ForEach-Object {
                if ($_.IsSeparator) { 'S:' + $_.Value + '|' }
                else { 'T:' + $_.DetectedTypeId + '|' }
            }) -join '')
}

<#
.SYNOPSIS
    Infers common field data types across multiple files sharing the same structural pattern.

.DESCRIPTION
    Evaluates candidate data types across a collection of tokens from the same position across multiple files.
    Determines if the position consistently represents a specific type (e.g., DateTime) or requires user resolution.

.PARAMETER Tokens
    Collection of value tokens at a specific field index from all analyzed files.

.OUTPUTS
    [PSCustomObject] Field inference result containing DetectedTypeId, CandidateTypes, IsAmbiguous, and Reason.
#>
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

<#
.SYNOPSIS
    Validates whether a string value conforms to a specified data type and optional format.

.DESCRIPTION
    Tests if a value matches a requested type ID ('Integer', 'Decimal', 'DateTime', 'Guid', 'Version', or custom type)
    and optional date format.

.PARAMETER Value
    The field value string to test.

.PARAMETER TypeId
    Target data type ID to test against.

.PARAMETER Format
    Optional required date format string if TypeId is 'DateTime'.

.PARAMETER CustomTypeRules
    Array of active custom type rules.

.OUTPUTS
    [Boolean] True if the value matches the type requirement, otherwise False.
#>
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

<#
.SYNOPSIS
    Validates a filename against a selected structural pattern and extracts field values.

.DESCRIPTION
    Enforces strict structural matching of token counts, separator exact literals, and field type constraints.
    Throws descriptive FNT pattern mismatch exceptions if validation fails.

.PARAMETER Name
    The filename string to test and parse.

.PARAMETER PatternTokens
    Array of template tokens defining the expected structure.

.PARAMETER TokenizerPattern
    Tokenizer regex pattern.

.PARAMETER FieldTypes
    Hashtable mapping field indices to expected type definitions.

.PARAMETER CustomTypeRules
    Array of active custom type rules.

.OUTPUTS
    [PSCustomObject] Object with Values hashtable (field index => extracted value) and actual Tokens array.
#>
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

<#
.SYNOPSIS
    Performs a deep clone of a PowerShell object via JSON serialization.

.DESCRIPTION
    Serializes the input object to JSON and deserializes it back to construct an isolated deep copy.

.PARAMETER InputObject
    The object to clone.

.OUTPUTS
    Cloned object copy or null if InputObject was null.
#>
function Copy-FNTObject {
    param([AllowNull()]$InputObject)

    if ($null -eq $InputObject) { return $null }
    return ($InputObject | ConvertTo-Json -Depth 20 | ConvertFrom-Json)
}

<#
.SYNOPSIS
    Normalizes legacy localized field type display labels into standard type IDs.

.DESCRIPTION
    Maps legacy localized string representations (e.g. 'Niejednoznaczny', 'Data', 'dziesiętna', 'Wersja')
    to standard internal type identifiers ('Ambiguous', 'DateTime', 'Decimal', 'Version', etc.).

.PARAMETER TypeLabel
    The raw legacy type label string.

.OUTPUTS
    [String] Standard type ID.
#>
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

<#
.SYNOPSIS
    Normalizes and upgrades application configuration objects to schema Version 2.

.DESCRIPTION
    Ensures that config object contains all required properties (Version = 2, Language = PL/EN/DE,
    Theme = Dark/Light, and CustomTypeRules array), supplying default values for missing entries.

.PARAMETER Config
    The raw configuration object loaded from config.json.

.OUTPUTS
    [PSCustomObject] Normalized configuration object.
#>
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
    if (-not $normalized.PSObject.Properties['Theme'] -or $normalized.Theme -notin @('Light', 'Dark')) {
        if ($normalized.PSObject.Properties['Theme']) { $normalized.Theme = 'Dark' }
        else { $normalized | Add-Member -NotePropertyName Theme -NotePropertyValue 'Dark' }
    }
    return $normalized
}

<#
.SYNOPSIS
    Updates the selected UI language in the configuration object.

.DESCRIPTION
    Sets the Language property of a normalized configuration object to 'PL', 'EN', or 'DE'.

.PARAMETER Config
    The target configuration object.

.PARAMETER Language
    The two-letter language code ('PL', 'EN', or 'DE').

.OUTPUTS
    [PSCustomObject] Updated configuration object.
#>
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

<#
.SYNOPSIS
    Updates the selected visual theme in the configuration object.

.DESCRIPTION
    Sets the Theme property of a normalized configuration object to 'Light' or 'Dark'.

.PARAMETER Config
    The target configuration object.

.PARAMETER Theme
    The visual theme name ('Light' or 'Dark').

.OUTPUTS
    [PSCustomObject] Updated configuration object.
#>
function Set-FNTConfigTheme {
    [CmdletBinding()]
    param(
        [AllowNull()]$Config,
        [Parameter(Mandatory)][ValidateSet('Light', 'Dark')][string]$Theme
    )

    $normalized = ConvertTo-FNTConfig $Config
    $normalized.Theme = $Theme
    return $normalized
}

<#
.SYNOPSIS
    Validates custom data type rules for correct syntax, ID constraints, and duplicate IDs.

.DESCRIPTION
    Checks custom regex rules loaded from configuration or UI tab 6. Verifies regex compilation,
    ID identifier syntax ('^[A-Za-z][A-Za-z0-9_.-]*$'), and checks for duplicate IDs.
    Returns separate arrays of ValidRules and Error objects.

.PARAMETER Rules
    Array of rule objects to validate.

.OUTPUTS
    [PSCustomObject] Object containing ValidRules array and Errors array.
#>
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

<#
.SYNOPSIS
    Normalizes profile objects to schema Version 2.

.DESCRIPTION
    Upgrades legacy profile structures by populating missing property fields (SchemaVersion = 2,
    TokenRegex, Fields, Mappings, OutputParts, KeepExtension, NewExtension, Transforms, IsVirtual, etc.).

.PARAMETER Profile
    Raw profile object read from JSON file.

.OUTPUTS
    [PSCustomObject] Normalized profile object.
#>
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

<#
.SYNOPSIS
    Extracts comprehensive filesystem, COM shell extended properties, EXIF image attributes, Office document properties, and MD5/SHA256 hashes.

.DESCRIPTION
    Reads file properties including creation time, last write time, COM Shell properties (Author, Title, Audio tags),
    Bitmap EXIF data (DateTaken, Dimensions, Camera model), Office OpenXML core properties (Creator, Title, Subject),
    and calculates MD5 and SHA256 checksums.

.PARAMETER Path
    Absolute or relative path to the target file.

.OUTPUTS
    [PSCustomObject] Ordered dictionary containing extracted file metadata fields.
#>
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
                    $propertyAliases = @{
                        Author = @('author', 'authors', 'autor', 'autoren', 'autorzy')
                        Title  = @('title', 'titel', 'tytul')
                        Artist = @('contributingartists', 'artists', 'interpretingartists', 'mitwirkendeinterpreten', 'wykonawcy')
                        Album  = @('album')
                        Year   = @('year', 'jahr', 'rok')
                    }
                    $shellValues = @{}
                    $folderKey = $parent.ToLowerInvariant()
                    if ($script:ShellMetadataHeaderIndexes.ContainsKey($folderKey)) {
                        $headerIndexes = $script:ShellMetadataHeaderIndexes[$folderKey]
                    }
                    else {
                        $headerIndexes = @{}
                        foreach ($index in 0..400) {
                            $header = [string]$folder.GetDetailsOf($folder.Items, $index)
                            if ([string]::IsNullOrWhiteSpace($header)) { continue }
                            $normalizedHeader = ($header.ToLowerInvariant() -replace '[^\p{L}\p{N}]', '')
                            foreach ($propertyName in $propertyAliases.Keys) {
                                if ($normalizedHeader -in $propertyAliases[$propertyName]) {
                                    $headerIndexes[$propertyName] = $index
                                    break
                                }
                            }
                        }
                        $script:ShellMetadataHeaderIndexes[$folderKey] = $headerIndexes
                    }
                    foreach ($propertyName in $headerIndexes.Keys) {
                        $value = [string]$folder.GetDetailsOf($item, $headerIndexes[$propertyName])
                        if (-not [string]::IsNullOrWhiteSpace($value)) {
                            $shellValues[$propertyName] = $value.Trim()
                        }
                    }

                    $author = if ($shellValues.Author) { $shellValues.Author } else { $folder.GetDetailsOf($item, 20) }
                    $title  = if ($shellValues.Title) { $shellValues.Title } else { $folder.GetDetailsOf($item, 21) }
                    if ($author) { $result.Author = [string]$author.Trim() }
                    if ($title)  { $result.Title  = [string]$title.Trim() }

                    # Audio extended tags
                    $artist = if ($shellValues.Artist) { $shellValues.Artist } else { $folder.GetDetailsOf($item, 13) }
                    $album  = if ($shellValues.Album) { $shellValues.Album } else { $folder.GetDetailsOf($item, 14) }
                    $year   = if ($shellValues.Year) { $shellValues.Year } else { $folder.GetDetailsOf($item, 15) }
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

        if (-not $result.Author -and $result.DocCreator) {
            $result.Author = $result.DocCreator
        }

        if ($result.Author) {
            $authorClean = $result.Author -replace '[^\p{L}\s,-]', ''
            $parts = @(@($authorClean -split '\s+|,') | Where-Object { $_.Trim() })
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

<#
.SYNOPSIS
    Normalizes a raw author string into an 8-character compliance segment (Surname7 + 1 Initial).

.DESCRIPTION
    Parses full name strings (SurnameFirst, GivenFirst, or concatenated) and formats them into
    a standard 8-character string consisting of up to 7 Titlecase surname characters and 1 Uppercase initial character,
    padded on the right with hyphens '-' if shorter.

.PARAMETER Raw
    Raw author name string.

.OUTPUTS
    [String] Normalized 8-character author segment string (e.g. 'KowalskJ' or 'MnichA---').
#>
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

    if ($cleanNoHyphens -cmatch '^(?<Surname>\p{Lu}\p{Ll}+)(?<Given>\p{Lu}\p{Ll}*)$') {
        # Concatenated full name: "KowalskiJan" -> "KowalskJ"
        $surname = $Matches.Surname
        $initial = $Matches.Given.Substring(0, 1)
    }
    elseif ($isLastUpper -and ($cleanNoHyphens.Length -eq 2 -or $cleanNoHyphens.Substring(0, $cleanNoHyphens.Length - 1) -cmatch '[\p{Ll}]')) {
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

<#
.SYNOPSIS
    Validates a filename against the enterprise naming scheme YYYYMMDD_XxxxxxxY_Zzzzz_Freitext.

.DESCRIPTION
    Checks if a filename base complies with the required structure: 8-digit date (yyyyMMdd), 8-character author segment,
    document type, free text description, and optional version suffix (v1, v2). Identifies specific violations.

.PARAMETER BaseName
    Filename string without extension.

.OUTPUTS
    [PSCustomObject] Object containing IsCompliant boolean, Segments hashtable, Violations array, and AuthorAnalysis object.
#>
function Test-FNTNamingConvention {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$BaseName
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

Export-ModuleMember -Function Get-FNTLexicalTokens, Get-FNTTypeCandidates, Get-FNTTokens, Get-FNTTokenSignature, Get-FNTFieldInference, Test-FNTValueType, Match-FNTNamePattern, ConvertFrom-FNTLegacyTypeLabel, ConvertTo-FNTConfig, Set-FNTConfigLanguage, Set-FNTConfigTheme, Test-FNTCustomTypeRules, ConvertTo-FNTProfile, Get-FNTFileMetadata, Test-FNTNamingConvention, Get-FNTNormalizedAuthorSegment