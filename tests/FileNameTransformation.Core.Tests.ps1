BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..\FileNameTransformation.Core.psm1') -Force
    $defaultPattern = '(?<value>[^_\-\s]+)|(?<sep>[_\-\s]+)'
}

Describe 'Get-FNTLexicalTokens' {
    It 'rejects an invalid tokenizer regex' {
        { Get-FNTLexicalTokens -Name 'Report_1' -Pattern '[invalid' } |
            Should -Throw '*Invalid tokenizer regex*'
    }

    It 'requires the separator capture group' {
        { Get-FNTLexicalTokens -Name 'Report' -Pattern '(?<value>.+)' } |
            Should -Throw "*named group 'sep'*"
    }

    It 'rejects zero-length matches' {
        { Get-FNTLexicalTokens -Name 'Report' -Pattern '(?<sep>_*)|(?<value>.+)' } |
            Should -Throw '*zero-length match*'
    }

    It 'rejects unmatched filename characters' {
        { Get-FNTLexicalTokens -Name 'Report_1' -Pattern '(?<value>[A-Za-z]+)|(?<sep>_)' } |
            Should -Throw '*position 7*'
    }

    It 'attaches a stable diagnostic code and position' {
        try {
            Get-FNTLexicalTokens -Name 'Report_1' -Pattern '(?<value>[A-Za-z]+)|(?<sep>_)'
            throw 'Expected tokenizer failure.'
        }
        catch {
            $_.Exception.Data['FNTCode'] | Should -Be 'Tokenizer.IncompleteCoverage'
            $_.Exception.Data['Position'] | Should -Be 7
        }
    }
}

Describe 'Get-FNTTypeCandidates' {
    It 'detects <Value> as <TypeId>' -TestCases @(
        @{ Value = '123456'; TypeId = 'Integer' }
        @{ Value = '123.45'; TypeId = 'Decimal' }
        @{ Value = '1.2.3'; TypeId = 'Version' }
        @{ Value = '6F9619FF-8B86-D011-B42D-00C04FC964FF'; TypeId = 'Guid' }
        @{ Value = '2025-01-16'; TypeId = 'DateTime' }
    ) {
        param($Value, $TypeId)

        $candidateIds = @(Get-FNTTypeCandidates -Value $Value | ForEach-Object { $_.TypeId })
        $candidateIds | Should -Contain $TypeId
    }

    It 'retains competing date and integer interpretations' {
        $candidates = @(Get-FNTTypeCandidates -Value '010225')

        @($candidates.TypeId) | Should -Contain 'Integer'
        @($candidates.TypeId) | Should -Contain 'DateTime'
        $candidates.Count | Should -BeGreaterThan 1
    }

    It 'does not accept an invalid calendar date' {
        $candidateIds = @(Get-FNTTypeCandidates -Value '2025-02-31' | ForEach-Object { $_.TypeId })

        $candidateIds | Should -Not -Contain 'DateTime'
    }
}

Describe 'Get-FNTTokens' {
    It 'merges a separated date into one typed field' {
        $tokens = @(Get-FNTTokens -Name 'Report_2025-01-16_Final' -Pattern $defaultPattern)

        $tokens.Count | Should -Be 5
        $tokens[2].Value | Should -Be '2025-01-16'
        $tokens[2].DetectedTypeId | Should -Be 'DateTime'
        $tokens[2].LexicalParts.Count | Should -Be 5
    }

    It 'merges a punctuated GUID into one typed field' {
        $guid = '6F9619FF-8B86-D011-B42D-00C04FC964FF'
        $tokens = @(Get-FNTTokens -Name "ID_${guid}_Final" -Pattern $defaultPattern)

        $tokens[2].Value | Should -Be $guid
        $tokens[2].DetectedTypeId | Should -Be 'Guid'
    }

    It 'marks compact numeric dates as ambiguous' {
        $tokens = @(Get-FNTTokens -Name 'Report_010225_Final' -Pattern $defaultPattern)
        $token = $tokens[2]

        $token.DetectedTypeId | Should -Be 'Ambiguous'
        $token.IsAmbiguous | Should -BeTrue
        $token.CandidateIds | Should -Contain 'Integer'
        $token.CandidateIds | Should -Contain 'DateTime'
    }

    It 'merges an enabled custom composite type' {
        $rules = @(
            [pscustomobject]@{
                Id             = 'DepartmentCode'
                Pattern        = '^[A-Z]{3}-\d{3}$'
                Enabled        = $true
                AllowComposite = $true
            }
        )

        $tokens = @(Get-FNTTokens -Name 'File_ABC-123_Final' -Pattern $defaultPattern -CustomTypeRules $rules)

        $tokens[2].Value | Should -Be 'ABC-123'
        $tokens[2].DetectedTypeId | Should -Be 'Custom:DepartmentCode'
    }

    It 'keeps ordinary separators as structural tokens' {
        $tokens = @(Get-FNTTokens -Name 'Report_Alpha_Final' -Pattern $defaultPattern)

        @($tokens | Where-Object IsSeparator).Count | Should -Be 2
        $tokens[2].DetectedTypeId | Should -Be 'Text'
    }
}

Describe 'Get-FNTFieldInference' {
    It 'infers one date format from every sample in the field' {
        $first = @(Get-FNTTokens -Name '2025-01-16' -Pattern $defaultPattern)[0]
        $second = @(Get-FNTTokens -Name '2025-02-17' -Pattern $defaultPattern)[0]

        $inference = Get-FNTFieldInference -Tokens @($first, $second)

        $inference.DetectedTypeId | Should -Be 'DateTime'
        $inference.CandidateTypes[0].Format | Should -Be 'yyyy-MM-dd'
        $inference.IsAmbiguous | Should -BeFalse
    }

    It 'retains ambiguity shared by every compact numeric sample' {
        $first = @(Get-FNTTokens -Name '010225' -Pattern $defaultPattern)[0]
        $second = @(Get-FNTTokens -Name '020225' -Pattern $defaultPattern)[0]

        $inference = Get-FNTFieldInference -Tokens @($first, $second)

        $inference.DetectedTypeId | Should -Be 'Ambiguous'
        @($inference.CandidateTypes.TypeId) | Should -Contain 'Integer'
        @($inference.CandidateTypes.TypeId) | Should -Contain 'DateTime'
        $inference.IsAmbiguous | Should -BeTrue
    }
}

Describe 'Match-FNTNamePattern' {
    BeforeEach {
        $patternTokens = @(Get-FNTTokens -Name 'Report_100_Final' -Pattern $defaultPattern)
    }

    It 'extracts values from the same strict structure' {
        $match = Match-FNTNamePattern -Name 'Invoice_200_Ready' -PatternTokens $patternTokens -TokenizerPattern $defaultPattern

        $match.Values[0] | Should -Be 'Invoice'
        $match.Values[2] | Should -Be '200'
        $match.Values[4] | Should -Be 'Ready'
    }

    It 'rejects a changed separator with its token position' {
        { Match-FNTNamePattern -Name 'Invoice-200_Ready' -PatternTokens $patternTokens -TokenizerPattern $defaultPattern } |
            Should -Throw "*token 2, offset 7*expected separator '_'*found '-'*"
    }

    It 'rejects a missing structural segment' {
        { Match-FNTNamePattern -Name 'Invoice_200' -PatternTokens $patternTokens -TokenizerPattern $defaultPattern } |
            Should -Throw '*expected 5 tokens, found 3*'
    }

    It 'rejects a value that violates the resolved field type' {
        $fieldTypes = @{ 2 = [pscustomobject]@{ TypeId = 'Integer'; Format = $null } }

        { Match-FNTNamePattern -Name 'Invoice_ABC_Ready' -PatternTokens $patternTokens -TokenizerPattern $defaultPattern -FieldTypes $fieldTypes } |
            Should -Throw "*token 3, offset 8*'ABC'*not type 'Integer'*"
    }
}

Describe 'Configuration compatibility' {
    It 'adds versioned defaults to a legacy language-only configuration' {
        $config = ConvertTo-FNTConfig ([pscustomobject]@{ Language = 'EN' })

        $config.Version | Should -Be 2
        $config.Language | Should -Be 'EN'
        $config.Theme | Should -Be 'Dark'
        @($config.CustomTypeRules).Count | Should -Be 0
    }

    It 'supports changing theme setting cleanly' {
        $source = [pscustomobject]@{
            Version  = 2
            Language = 'EN'
            Theme    = 'Light'
        }

        $updated = Set-FNTConfigTheme -Config $source -Theme 'Dark'

        $updated.Theme | Should -Be 'Dark'
        $source.Theme | Should -Be 'Light'
    }

    It 'preserves custom rules and unknown settings when changing language' {
        $source = [pscustomobject]@{
            Version         = 2
            Language        = 'EN'
            CustomTypeRules = @([pscustomobject]@{ Id = 'Code'; Pattern = '^X\d+$' })
            FutureSetting   = 'keep-me'
        }

        $updated = Set-FNTConfigLanguage -Config $source -Language 'DE'

        $updated.Language | Should -Be 'DE'
        $updated.CustomTypeRules[0].Id | Should -Be 'Code'
        $updated.FutureSetting | Should -Be 'keep-me'
        $source.Language | Should -Be 'EN'
    }

    It 'preserves custom rules through a temporary configuration file update' {
        $path = Join-Path $TestDrive 'config.json'
        [pscustomobject]@{
            Version         = 2
            Language        = 'PL'
            CustomTypeRules = @([pscustomobject]@{ Id = 'Case'; Pattern = '^C-\d+$'; Enabled = $true })
        } | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $path -Encoding UTF8

        $loaded = Get-Content -LiteralPath $path -Raw -Encoding UTF8 | ConvertFrom-Json
        $updated = Set-FNTConfigLanguage -Config $loaded -Language 'EN'
        $updated | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $path -Encoding UTF8
        $saved = Get-Content -LiteralPath $path -Raw -Encoding UTF8 | ConvertFrom-Json

        $saved.Language | Should -Be 'EN'
        $saved.CustomTypeRules[0].Id | Should -Be 'Case'
        $saved.CustomTypeRules[0].Pattern | Should -Be '^C-\d+$'
    }

    It 'produces the same structural signature after a language change' {
        $source = [pscustomobject]@{ Language = 'PL'; CustomTypeRules = @() }
        $polishConfig = ConvertTo-FNTConfig $source
        $germanConfig = Set-FNTConfigLanguage -Config $polishConfig -Language 'DE'

        $polishTokens = Get-FNTTokens -Name 'Report_2025-01-16_Final' -Pattern $defaultPattern -CustomTypeRules $polishConfig.CustomTypeRules
        $germanTokens = Get-FNTTokens -Name 'Report_2025-01-16_Final' -Pattern $defaultPattern -CustomTypeRules $germanConfig.CustomTypeRules

        Get-FNTTokenSignature $polishTokens | Should -Be (Get-FNTTokenSignature $germanTokens)
    }
}

Describe 'Custom type rule validation' {
    It 'normalizes valid rules and supplies defaults' {
        $result = Test-FNTCustomTypeRules @([pscustomobject]@{
                Id      = 'Department.Code'
                Pattern = '^[A-Z]{3}-\d{3}$'
            })

        $result.Errors.Count | Should -Be 0
        $result.ValidRules.Count | Should -Be 1
        $result.ValidRules[0].DisplayName | Should -Be 'Department.Code'
        $result.ValidRules[0].Enabled | Should -BeTrue
        $result.ValidRules[0].AllowComposite | Should -BeFalse
    }

    It 'reports invalid IDs, duplicate IDs, empty patterns, and invalid regexes separately' {
        $result = Test-FNTCustomTypeRules @(
            [pscustomobject]@{ Id = 'Valid'; Pattern = '^A$' }
            [pscustomobject]@{ Id = 'valid'; Pattern = '^B$' }
            [pscustomobject]@{ Id = '123bad'; Pattern = '^C$' }
            [pscustomobject]@{ Id = 'Empty'; Pattern = '' }
            [pscustomobject]@{ Id = 'Broken'; Pattern = '[invalid' }
        )

        $result.ValidRules.Count | Should -Be 1
        $result.Errors.Count | Should -Be 4
        @($result.Errors.Message) -join ' ' | Should -Match 'duplicates ID'
        @($result.Errors.Message) -join ' ' | Should -Match 'invalid ID'
        @($result.Errors.Message) -join ' ' | Should -Match 'empty pattern'
        @($result.Errors.Message) -join ' ' | Should -Match 'invalid regex'
    }

    It 'keeps disabled valid rules but excludes them from recognition' {
        $validation = Test-FNTCustomTypeRules @([pscustomobject]@{
                Id          = 'DisabledCode'
                DisplayName = 'Disabled code'
                Pattern     = '^ABC$'
                Enabled     = $false
            })

        $validation.ValidRules.Count | Should -Be 1
        $validation.ValidRules[0].Enabled | Should -BeFalse
        @(Get-FNTTypeCandidates -Value 'ABC' -CustomTypeRules $validation.ValidRules).Count | Should -Be 0
    }

    It 'uses the configured display name on a custom candidate' {
        $validation = Test-FNTCustomTypeRules @([pscustomobject]@{
                Id          = 'DepartmentCode'
                DisplayName = 'Department code'
                Pattern     = '^[A-Z]{3}-\d{3}$'
                Enabled     = $true
            })

        $candidate = @(Get-FNTTypeCandidates -Value 'ABC-123' -CustomTypeRules $validation.ValidRules |
            Where-Object TypeId -eq 'Custom:DepartmentCode')[0]

        $candidate.DisplayName | Should -Be 'Department code'
    }
}

Describe 'Profile compatibility' {
    It 'maps localized legacy type labels to stable IDs' -TestCases @(
        @{ Label = 'Data (yyyyMMdd)'; Expected = 'DateTime' }
        @{ Label = 'Date (yyyy-MM-dd)'; Expected = 'DateTime' }
        @{ Label = 'Datum (dd-MM-yyyy)'; Expected = 'DateTime' }
        @{ Label = 'Liczba'; Expected = 'Integer' }
        @{ Label = 'Number'; Expected = 'Integer' }
        @{ Label = 'Zahl'; Expected = 'Integer' }
        @{ Label = 'Tekst'; Expected = 'Text' }
    ) {
        param($Label, $Expected)

        ConvertFrom-FNTLegacyTypeLabel $Label | Should -Be $Expected
    }

    It 'migrates a legacy profile while preserving workflow data' {
        $legacy = [pscustomobject]@{
            Name          = 'Legacy'
            Fields        = @([pscustomobject]@{
                    Index        = 2
                    Name         = 'OrderId'
                    DetectedType = 'Number'
                    Transforms   = @([pscustomobject]@{ Type = 'Pad'; Width = 8 })
                })
            Mappings      = @([pscustomobject]@{
                    Name        = 'Orders'
                    InputField  = 'OrderId'
                    OutputField = 'Customer'
                })
            OutputParts   = @([pscustomobject]@{ Type = 'Field'; Value = 'Customer' })
            KeepExtension = $true
            NewExtension  = '.csv'
        }

        $profile = ConvertTo-FNTProfile $legacy

        $profile.SchemaVersion | Should -Be 2
        $profile.Fields[0].PartIndex | Should -Be 2
        $profile.Fields[0].DetectedTypeId | Should -Be 'Integer'
        $profile.Fields[0].SelectedTypeId | Should -Be 'Auto'
        $profile.Fields[0].Transforms[0].Type | Should -Be 'Pad'
        $profile.Mappings[0].InputField | Should -Be 'OrderId'
        $profile.OutputParts[0].Value | Should -Be 'Customer'
        $profile.TokenRegex | Should -Not -BeNullOrEmpty
    }

    It 'round trips a version 2 profile without losing typed settings' {
        $source = [pscustomobject]@{
            SchemaVersion = 2
            Name          = 'Typed'
            TokenRegex    = '(?<value>[^_]+)|(?<sep>_)'
            Fields        = @([pscustomobject]@{
                    PartIndex       = 2
                    DisplayIndex    = '2'
                    Name            = 'Created'
                    DetectedTypeId  = 'Ambiguous'
                    CandidateTypes  = @(
                        [pscustomobject]@{ TypeId = 'Integer'; Format = $null }
                        [pscustomobject]@{ TypeId = 'DateTime'; Format = 'yyMMdd' }
                    )
                    IsAmbiguous     = $true
                    SelectedTypeId  = 'DateTime'
                    SelectedFormat  = 'yyMMdd'
                    IsVirtual       = $false
                    Transforms      = @([pscustomobject]@{ Type = 'DateFormat'; InputFormat = 'yyMMdd'; OutputFormat = 'yyyy-MM-dd' })
                })
            Mappings      = @()
            OutputParts   = @([pscustomobject]@{ Type = 'Field'; Value = 'Created' })
            KeepExtension = $false
            NewExtension  = 'txt'
        }
        $jsonRoundTrip = $source | ConvertTo-Json -Depth 20 | ConvertFrom-Json

        $profile = ConvertTo-FNTProfile $jsonRoundTrip

        $profile.TokenRegex | Should -Be $source.TokenRegex
        $profile.Fields[0].SelectedTypeId | Should -Be 'DateTime'
        $profile.Fields[0].SelectedFormat | Should -Be 'yyMMdd'
        @($profile.Fields[0].CandidateTypes).Count | Should -Be 2
        $profile.Fields[0].Transforms[0].OutputFormat | Should -Be 'yyyy-MM-dd'
        $profile.OutputParts[0].Value | Should -Be 'Created'
        $profile.KeepExtension | Should -BeFalse
        $profile.NewExtension | Should -Be 'txt'
    }

    It 'round trips a version 2 profile through a temporary JSON file' {
        $path = Join-Path $TestDrive 'Typed.json'
        $source = [pscustomobject]@{
            Name          = 'Typed file'
            Fields        = @([pscustomobject]@{
                    PartIndex      = 0
                    Name           = 'OrderId'
                    DetectedTypeId = 'Integer'
                    SelectedTypeId = 'Integer'
                    Transforms     = @([pscustomobject]@{ Type = 'Pad'; Width = 8 })
                })
            Mappings      = @([pscustomobject]@{ Name = 'Orders'; InputField = 'OrderId'; OutputField = 'Customer' })
            OutputParts   = @([pscustomobject]@{ Type = 'Field'; Value = 'Customer' })
            KeepExtension = $true
            NewExtension  = ''
        }
        ConvertTo-FNTProfile $source | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $path -Encoding UTF8

        $saved = Get-Content -LiteralPath $path -Raw -Encoding UTF8 | ConvertFrom-Json
        $profile = ConvertTo-FNTProfile $saved

        $profile.SchemaVersion | Should -Be 2
        $profile.Fields[0].SelectedTypeId | Should -Be 'Integer'
        $profile.Fields[0].Transforms[0].Width | Should -Be 8
        $profile.Mappings[0].OutputField | Should -Be 'Customer'
        $profile.OutputParts[0].Value | Should -Be 'Customer'
    }
}

Describe 'Test-FNTNamingConvention' {
    It 'classifies an empty basename as noncompliant' {
        $res = Test-FNTNamingConvention -BaseName ''

        $res.IsCompliant | Should -BeFalse
        $res.Segments.FreeText | Should -Be ''
        $res.Violations | Should -Contain 'Compliance_StructureBad'
        $res.Violations | Should -Contain 'Compliance_DateBad'
    }

    It 'validates compliant filename' {
        $res = Test-FNTNamingConvention -BaseName '20260721_MusteraM_Protokoll_Neubau-LBO_v1'
        $res.IsCompliant | Should -Be $true
        $res.Segments.Date | Should -Be '20260721'
        $res.Segments.Author | Should -Be 'MusteraM'
        $res.Segments.DocType | Should -Be 'Protokoll'
        $res.Segments.FreeText | Should -Be 'Neubau-LBO'
        $res.Segments.Version | Should -Be 'v1'
        $res.AuthorAnalysis.DetectedFormat | Should -Be 'SurnameFirst'
    }

    It 'validates compliant filename with hyphen padding' {
        $res1 = Test-FNTNamingConvention -BaseName '20260721_MajkJ---_Protokoll_Neubau-LBO_v1'
        $res1.IsCompliant | Should -Be $true
        $res1.Segments.Author | Should -Be 'MajkJ---'
        $res1.AuthorAnalysis.DetectedFormat | Should -Be 'SurnameFirst'

        $res2 = Test-FNTNamingConvention -BaseName '20260721_MnichA--_Protokoll_Neubau-LBO_v1'
        $res2.IsCompliant | Should -Be $true
        $res2.Segments.Author | Should -Be 'MnichA--'
    }

    It 'detects reversed author with hyphens (GivenFirst)' {
        $res = Test-FNTNamingConvention -BaseName '20260721_JMajk---_Protokoll_Neubau-LBO_v1'
        $res.IsCompliant | Should -Be $false
        $res.AuthorAnalysis.DetectedFormat | Should -Be 'GivenFirst'
        $res.AuthorAnalysis.SuggestedAuthor | Should -Be 'MajkJ---'
        $res.Violations | Should -Contain 'Compliance_AuthorReversed'
    }

    It 'detects short author segment and suggests hyphens' {
        $res = Test-FNTNamingConvention -BaseName '20260721_MajkJ_Protokoll_Neubau-LBO_v1'
        $res.IsCompliant | Should -Be $false
        $res.AuthorAnalysis.DetectedFormat | Should -Be 'TooShort'
        $res.AuthorAnalysis.SuggestedAuthor | Should -Be 'MajkJ---'
        $res.Violations | Should -Contain 'Compliance_AuthorShort'
    }

    It 'detects short author segment with single hyphen and pads to 8 chars' {
        $res = Test-FNTNamingConvention -BaseName '20201221_MnichA-_SerieNumeracji_2021'
        $res.IsCompliant | Should -Be $false
        $res.AuthorAnalysis.DetectedFormat | Should -Be 'TooShort'
        $res.AuthorAnalysis.SuggestedAuthor | Should -Be 'MnichA--'
        $res.Violations | Should -Contain 'Compliance_AuthorShort'
    }

    It 'parses 3-part filename with Date, Author, and FreeText (missing DocType)' {
        $res = Test-FNTNamingConvention -BaseName '20211104_MnichA-_Kilometry SK627GE 10.2021'
        $res.IsCompliant | Should -Be $false
        $res.Segments.Date | Should -Be '20211104'
        $res.Segments.Author | Should -Be 'MnichA-'
        $res.Segments.DocType | Should -Be ''
        $res.Segments.FreeText | Should -Be 'Kilometry SK627GE 10.2021'
        $res.AuthorAnalysis.SuggestedAuthor | Should -Be 'MnichA--'
        $res.Violations | Should -Contain 'Compliance_StructureBad'
        $res.Violations | Should -Contain 'Compliance_AuthorShort'
    }

    It 'detects invalid date format' {
        $res = Test-FNTNamingConvention -BaseName '2026-07-21_MusteraM_Protokoll_Neubau-LBO_v1'
        $res.IsCompliant | Should -Be $false
        $res.Violations | Should -Contain 'Compliance_DateBad'
    }
}

Describe 'Get-FNTNormalizedAuthorSegment' {
    It 'normalizes various author input strings to exactly 8 characters' -TestCases @(
        @{ Raw = 'MnichA-';      Expected = 'MnichA--' }
        @{ Raw = 'MajkJ';       Expected = 'MajkJ---' }
        @{ Raw = 'AMnich';      Expected = 'MnichA--' }
        @{ Raw = 'AMnich-';     Expected = 'MnichA--' }
        @{ Raw = 'JMajk---';    Expected = 'MajkJ---' }
        @{ Raw = 'MusteraM';    Expected = 'MusteraM' }
        @{ Raw = 'MustermannM'; Expected = 'MustermM' }
        @{ Raw = 'KowalskiJan'; Expected = 'KowalskJ' }
    ) {
        param($Raw, $Expected)
        Get-FNTNormalizedAuthorSegment $Raw | Should -Be $Expected
    }
}

Describe 'Field transformations' {
    BeforeAll {
        function global:T([string]$Key) { $Key }
        . (Join-Path $PSScriptRoot '..\src\Transforms.ps1')
    }

    It 'executes PowerShell expressions with the current value in Windows PowerShell' {
        $transform = [pscustomobject]@{
            Type       = 'Expression'
            Expression = 'if ($_ -match ''^\s*([^,]+),\s*(.)'') { "$($matches[1])$($matches[2])" } else { $_ }'
        }

        ApplyTransforms 'Kowalski, Jan' @($transform) | Should -Be 'KowalskiJ'
    }
}

Describe 'Phase 2: Extended Metadata & Dictionary Readers' {
    BeforeAll {
        . (Join-Path $PSScriptRoot '..\src\Mappings.ps1')
    }

    It 'returns extended metadata schema from Get-FNTFileMetadata' {
        $testFile = [System.IO.Path]::GetTempFileName()
        try {
            Set-Content -LiteralPath $testFile -Value "Test Content" -Encoding UTF8
            $meta = Get-FNTFileMetadata -Path $testFile
            $meta.PSObject.Properties['CreationDateStr'] | Should -Not -BeNullOrEmpty
            $meta.PSObject.Properties['HashMD5'] | Should -Not -BeNullOrEmpty
            $meta.HashMD5.Length | Should -Be 32
            $meta.HashSHA256.Length | Should -Be 64
        }
        finally {
            Remove-Item $testFile -ErrorAction SilentlyContinue
        }
    }

    It 'skips hash calculation when requested' {
        $testFile = [System.IO.Path]::GetTempFileName()
        try {
            Set-Content -LiteralPath $testFile -Value 'Test Content' -Encoding UTF8
            $meta = Get-FNTFileMetadata -Path $testFile -SkipHashes

            $meta.CreationDateStr | Should -Not -BeNullOrEmpty
            $meta.HashMD5 | Should -BeNullOrEmpty
            $meta.HashSHA256 | Should -BeNullOrEmpty
        }
        finally {
            Remove-Item $testFile -ErrorAction SilentlyContinue
        }
    }

    It 'reads dictionary headers from JSON file' {
        $jsonFile = [System.IO.Path]::GetTempFileName() + '.json'
        try {
            $jsonContent = '[{"DepartmentCode":"IT", "DepartmentName":"Information Technology"}]'
            Set-Content -LiteralPath $jsonFile -Value $jsonContent -Encoding UTF8
            $headers = Get-FNTDictionaryHeaders $jsonFile
            $headers.Format | Should -Be 'JSON'
            $headers.Headers | Should -Contain 'DepartmentCode'
            $headers.Headers | Should -Contain 'DepartmentName'
        }
        finally {
            Remove-Item $jsonFile -ErrorAction SilentlyContinue
        }
    }

    It 'removes CSV quoting from detected headers' {
        $csvFile = [System.IO.Path]::GetTempFileName() + '.csv'
        try {
            Set-Content -LiteralPath $csvFile -Value '"EmployeeID";"DisplayName"' -Encoding UTF8
            Add-Content -LiteralPath $csvFile -Value '"E1001";"Kowalski, Jan"' -Encoding UTF8

            $headers = Get-FNTDictionaryHeaders $csvFile

            $headers.Delimiter | Should -Be ';'
            $headers.Headers | Should -Contain 'EmployeeID'
            $headers.Headers | Should -Contain 'DisplayName'
            $headers.Headers | Should -Not -Contain '"EmployeeID"'
        }
        finally {
            Remove-Item $csvFile -ErrorAction SilentlyContinue
        }
    }

    It 'reads dictionary headers from XML file' {
        $xmlFile = [System.IO.Path]::GetTempFileName() + '.xml'
        try {
            $xmlContent = '<Root><Item Code="DEV" Title="Developer"/></Root>'
            Set-Content -LiteralPath $xmlFile -Value $xmlContent -Encoding UTF8
            $headers = Get-FNTDictionaryHeaders $xmlFile
            $headers.Format | Should -Be 'XML'
            $headers.Headers | Should -Contain 'Title'
            $headers.Headers | Should -Contain '@Code'
        }
        finally {
            Remove-Item $xmlFile -ErrorAction SilentlyContinue
        }
    }
}