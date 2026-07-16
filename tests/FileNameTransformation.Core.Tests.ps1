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
            Should -Throw "*token 2*expected separator '_'*found '-'*"
    }

    It 'rejects a missing structural segment' {
        { Match-FNTNamePattern -Name 'Invoice_200' -PatternTokens $patternTokens -TokenizerPattern $defaultPattern } |
            Should -Throw '*expected 5 tokens, found 3*'
    }

    It 'rejects a value that violates the resolved field type' {
        $fieldTypes = @{ 2 = [pscustomobject]@{ TypeId = 'Integer'; Format = $null } }

        { Match-FNTNamePattern -Name 'Invoice_ABC_Ready' -PatternTokens $patternTokens -TokenizerPattern $defaultPattern -FieldTypes $fieldTypes } |
            Should -Throw "*token 3*'ABC'*not type 'Integer'*"
    }
}