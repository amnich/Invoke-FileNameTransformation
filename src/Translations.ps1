# Translations.ps1 — Loads locale JSON files and defines the T() lookup function.
# Dot-sourced by the main script; operates in $script: scope.

$script:Translations = @{}
$localesPath = Join-Path $script:ScriptRoot 'locales'

foreach ($lang in @('PL', 'EN', 'DE')) {
    $jsonPath = Join-Path $localesPath "$($lang.ToLower()).json"
    if (Test-Path -LiteralPath $jsonPath -PathType Leaf) {
        $json = Get-Content -LiteralPath $jsonPath -Raw -Encoding UTF8 | ConvertFrom-Json
        $ht = @{}
        $json.PSObject.Properties | ForEach-Object { $ht[$_.Name] = $_.Value }
        $script:Translations[$lang] = $ht
    }
    else {
        Write-Warning "Missing locale file: $jsonPath"
        $script:Translations[$lang] = @{}
    }
}

<#
.SYNOPSIS
    Returns the localized translation string for a given key in the current language context.

.PARAMETER Key
    The translation string key (e.g. 'Err_NoFiles', 'Btn_Browse').

.OUTPUTS
    [String] Localized string or key itself if missing.
#>
function T([string]$Key) {
    if ($script:Translations[$script:CurrentLanguage].ContainsKey($Key)) {
        return $script:Translations[$script:CurrentLanguage][$Key]
    }
    return $Key
}
