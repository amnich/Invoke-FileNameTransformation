<#
.SYNOPSIS
    Compiles and signs a standalone, self-contained executable for File Name Transformer using ps2exe.

.DESCRIPTION
    The ps2exe packaging script embeds all application assets—including the core module (FileNameTransformation.Core.psm1),
    WPF interface definition (MainWindow.xaml encoded in Base64), locale translation dictionaries (en.json, pl.json, de.json),
    and all src/ helper modules—into a single self-contained PowerShell script payload.

    It signs the temporary script using an available Code Signing certificate, compiles it into a native STA Windows GUI executable
    via ps2exe (with icon integration), and performs secondary dual SHA-256 Authenticode signing on the resulting EXE binary via signtool.exe.

.PARAMETER DebugBuild
    When specified, creates Invoke-FileNameTransformation.debug.exe with a visible console window and ps2exe debug flags enabled.
    By default (without this switch), builds a production GUI binary (Invoke-FileNameTransformation.exe) with no console window.

.EXAMPLE
    .\ps2exe.ps1
    Compiles, embeds assets, and signs the production executable (Invoke-FileNameTransformation.exe).

.EXAMPLE
    .\ps2exe.ps1 -DebugBuild
    Compiles a debug executable with console logging support (Invoke-FileNameTransformation.debug.exe).

.INPUTS
    None. Reads files from the local directory (Invoke-FileNameTransformation.ps1, MainWindow.xaml, locales/*.json, src/*.ps1).

.OUTPUTS
    [System.IO.FileInfo] Compiled EXE binary (Invoke-FileNameTransformation.exe or .debug.exe).

.NOTES
    Build Requirements:
    - Installed 'ps2exe' PowerShell module (`Install-Module ps2exe`).
    - Valid Code Signing certificate in Cert:\CurrentUser\My (optional, gracefully skipped with warning if absent).
    - Windows SDK `signtool.exe` for Authenticode binary timestamping (optional).
    - Application icon file specified in iconFile parameter.
#>

[CmdletBinding()]
param(
        [switch]$DebugBuild
)

$sourcePath = Join-Path $PSScriptRoot 'Invoke-FileNameTransformation.ps1'
$coreModulePath = Join-Path $PSScriptRoot 'FileNameTransformation.Core.psm1'
$xamlPath = Join-Path $PSScriptRoot 'MainWindow.xaml'
$localeDirectory = Join-Path $PSScriptRoot 'locales'
$sourceDirectory = Join-Path $PSScriptRoot 'src'
$localeFiles = @('pl.json', 'en.json', 'de.json')
$sourceFiles = @('UI.ps1', 'Analysis.ps1', 'Transforms.ps1', 'Mappings.ps1', 'Preview.ps1', 'Profiles.ps1', 'Compliance.ps1', 'Events.ps1')

if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) {
        throw "Missing application script: $sourcePath"
}
if (-not (Test-Path -LiteralPath $coreModulePath -PathType Leaf)) {
        throw "Missing required core module: $coreModulePath"
}
if (-not (Test-Path -LiteralPath $xamlPath -PathType Leaf)) {
        throw "Missing UI template: $xamlPath"
}
foreach ($fileName in $localeFiles) {
        $filePath = Join-Path $localeDirectory $fileName
        if (-not (Test-Path -LiteralPath $filePath -PathType Leaf)) {
                throw "Missing locale file: $filePath"
        }
}
foreach ($fileName in $sourceFiles) {
        $filePath = Join-Path $sourceDirectory $fileName
        if (-not (Test-Path -LiteralPath $filePath -PathType Leaf)) {
                throw "Missing application source file: $filePath"
        }
}

$temporarySourcePath = Join-Path ([IO.Path]::GetTempPath()) ('Invoke-FileNameTransformation_{0}.ps1' -f [guid]::NewGuid().ToString('N'))

try {
        $sourceContent = (Get-Content -LiteralPath $sourcePath -Raw -Encoding UTF8) -replace "`r`n", "`n"
        $coreContent = (Get-Content -LiteralPath $coreModulePath -Raw -Encoding UTF8) -replace "`r`n", "`n"
        $xamlContent = Get-Content -LiteralPath $xamlPath -Raw -Encoding UTF8
        $localeContent = @{}
        foreach ($fileName in $localeFiles) {
                $localeContent[[IO.Path]::GetFileNameWithoutExtension($fileName).ToUpperInvariant()] =
                Get-Content -LiteralPath (Join-Path $localeDirectory $fileName) -Raw -Encoding UTF8
        }

        # Runtime dependencies are embedded into the temporary script before ps2exe compiles it.
        $isolatedHostPattern = '(?ms)\[CmdletBinding\(\)\]\s*param\(.*?^\).*?if \(-not \$IsolatedHost -and -not \$NonInteractive\) \{.*?\}\s*(?=\$ErrorActionPreference)'
        if ($sourceContent -notmatch $isolatedHostPattern) {
                throw 'Could not find the isolated-host launcher in the application script.'
        }
        $sourceContent = [regex]::Replace($sourceContent, $isolatedHostPattern, "[CmdletBinding()]`nparam()`n`n", 1)

        $embeddedCore = $coreContent -replace '(?m)^Set-StrictMode -Version 2\.0\n?', ''
        $embeddedCore = $embeddedCore -replace '(?m)^Export-ModuleMember .*\n?', ''
        $bootstrapPattern = '(?ms)^\$coreModulePath = Join-Path \$script:ScriptRoot ''FileNameTransformation\.Core\.psm1''\nif \(-not \(Test-Path -LiteralPath \$coreModulePath -PathType Leaf\)\) \{\n    throw "Missing core module: \$coreModulePath"\n\}\nImport-Module \$coreModulePath -Force -DisableNameChecking\n?'

        if ($sourceContent -notmatch $bootstrapPattern) {
                throw 'Could not find the core module import block in the application script.'
        }

        $replacement = [System.Text.RegularExpressions.MatchEvaluator] { param($match) "`n$embeddedCore`n" }
        $mergedContent = [regex]::Replace($sourceContent, $bootstrapPattern, $replacement, 1)

        $xamlPayload = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($xamlContent))
        $xamlPattern = '(?ms)# Read MainWindow\.xaml from disk\n\$xamlPath = Join-Path \$script:ScriptRoot ''MainWindow\.xaml''\ntry \{\n    if \(-not \(Test-Path -LiteralPath \$xamlPath -PathType Leaf\)\) \{\n        throw "Missing UI template: \$xamlPath"\n    \}\n    \$xamlTemplate = \[System\.IO\.File\]::ReadAllText\(\$xamlPath\)\n\}\ncatch \{\n    Throw-StartupFailure "reading UI template ''\$xamlPath''" \$_\.Exception\n\}'
        $embeddedXaml = "# MainWindow.xaml is embedded in the packaged executable.`n`$xamlTemplate = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('$xamlPayload'))"
        if ($mergedContent -notmatch $xamlPattern) {
                throw 'Could not find the MainWindow.xaml load block in the application script.'
        }
        $mergedContent = [regex]::Replace($mergedContent, $xamlPattern, [Text.RegularExpressions.MatchEvaluator] { param($match) $embeddedXaml }, 1)

        $localeTranslationBlocks = $localeContent.Keys | Sort-Object | ForEach-Object {
                $language = $_
                $translations = $localeContent[$language] | ConvertFrom-Json
                $translationEntries = $translations.PSObject.Properties | ForEach-Object {
                        $keyPayload = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($_.Name))
                        $valuePayload = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes([string]$_.Value))
                        "        ([Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('$keyPayload'))) = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('$valuePayload'))"
                }
                "    '$language' = @{`n$($translationEntries -join "`n")`n    }"
        }
        $embeddedTranslations = @"
# Locale JSON files are embedded in the packaged executable.
`$script:Translations = @{
$($localeTranslationBlocks -join "`n")
}

function T([string]`$Key) {
    if (`$script:Translations[`$script:CurrentLanguage].ContainsKey(`$Key)) {
        return `$script:Translations[`$script:CurrentLanguage][`$Key]
    }
    return `$Key
}
"@
        $translationsImport = ". (Join-Path `$script:ScriptRoot 'src\Translations.ps1')"
        if (-not $mergedContent.Contains($translationsImport)) {
                throw 'Could not find the translations source import in the application script.'
        }
        $mergedContent = $mergedContent.Replace($translationsImport, $embeddedTranslations)

        foreach ($fileName in $sourceFiles) {
                $sourceImport = ". (Join-Path `$script:ScriptRoot 'src\$fileName')"
                if (-not $mergedContent.Contains($sourceImport)) {
                        throw "Could not find source import: $sourceImport"
                }
                $embeddedSource = (Get-Content -LiteralPath (Join-Path $sourceDirectory $fileName) -Raw -Encoding UTF8) -replace "`r`n", "`n"
                $mergedContent = $mergedContent.Replace($sourceImport, $embeddedSource)
        }

        Set-Content -LiteralPath $temporarySourcePath -Value $mergedContent -Encoding UTF8

        $signingCertificate = Get-ChildItem cert:\CurrentUser\my -ErrorAction SilentlyContinue |
        Where-Object EnhancedKeyUsageList -match '1.3.6.1.5.5.7.3.3' |
        Where-Object Subject -like '*Mnich*' |
        Sort-Object NotAfter -Descending |
        Select-Object -First 1
        if ($null -ne $signingCertificate) {
                $sourceSignature = Set-AuthenticodeSignature -Certificate $signingCertificate -IncludeChain All `
                        -TimestampServer 'http://timestamp.sectigo.com' -Force -FilePath $temporarySourcePath
                if ($sourceSignature.Status -ne 'Valid') {
                        Write-Warning "Temporary source signing status: $($sourceSignature.Status)"
                }
        }
        else {
                Write-Warning 'No suitable code-signing certificate was found in the current user certificate store. Compilation will proceed without code signing.'
        }
        $inputPath = $temporarySourcePath

        $Version = "2.0"
        Import-Module ps2exe

        $exePath = if ($DebugBuild) {
                $sourcePath -replace '\.ps1$', '.debug.exe'
        }
        else {
                $sourcePath -replace '\.ps1$', '.exe'
        }
        $ps2exeParameters = @{
                inputFile  = $inputPath
                outputFile = $exePath
                iconFile   = 'D:\Skrypty\Mnich_Adam_Skrypty\BGH.ico'
                title      = 'File Name Transformation'
                version    = $Version
                STA        = $true
        }
        if (-not $DebugBuild) {
                $ps2exeParameters.noConsole = $true
                $ps2exeParameters.noError = $true
        }
        else {
                $ps2exeParameters.Debug = $true
        }
        Invoke-ps2exe @ps2exeParameters

        if ($null -ne $signingCertificate) {
                $signtool = "C:\Program Files (x86)\Windows Kits\10\bin\x64\signtool.exe"
                if (-not (Test-Path -LiteralPath $signtool -PathType Leaf)) {
                        $signtoolCmd = Get-Command signtool.exe -ErrorAction SilentlyContinue
                        if ($signtoolCmd) { $signtool = $signtoolCmd.Source }
                }
                if (Test-Path -LiteralPath $signtool -PathType Leaf) {
                        $thumbprint = $signingCertificate.Thumbprint
                        $timestampUrl = "http://timestamp.sectigo.com"
                        & $signtool sign /sha1 $thumbprint /tr $timestampUrl /td SHA256 /fd SHA256 $exePath
                }
        }
}
finally {
        if (Test-Path -LiteralPath $temporarySourcePath) {
                Remove-Item -LiteralPath $temporarySourcePath -Force
        }
}