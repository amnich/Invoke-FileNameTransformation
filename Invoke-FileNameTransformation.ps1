#requires -Version 5.1
<#
.SYNOPSIS
    Launches the File Name Transformer GUI for bulk renaming and copying.

.DESCRIPTION
    FileNameTransformer.GUI is a WPF-based PowerShell script that enables bulk renaming and copying of files.
    It parses source filenames using structural patterns, applies value mappings from CSV dictionaries,
    and performs text transformations such as casing changes, substrings, replacements, padding, and date formatting.
    The script then builds target filenames from user-defined templates and offers a preview grid to validate
    results before copying or moving files.

    Tokenizer regex: the expression must match every filename segment. Use the named group
    (?<sep>...) for literal separators; all other matches are treated as value fields. The default
    (?<value>[^_\-\s]+)|(?<sep>[_\-\s]+) splits values on underscores, hyphens, and whitespace.
    A second semantic pass recognizes integers, invariant decimals, exact dates/times, GUIDs,
    versions, and configured custom regex types. Recognized composite values can span lexical
    separators. Ambiguous values retain all candidates and require a field type selection.

    Features:
    - Multi-language UI (Polish, English, German).
    - Saveable and reusable profiles in JSON format.
    - Pattern-based source file parsing.
    - Typed field inference with explicit ambiguity resolution.
    - Strict structural matching when applying a selected pattern to other files.
    - External CSV-based mapping support.
    - Preview and validation for collisions, missing values, and invalid characters.
    - Copy or move execution mode with audit logging.
    - Independent options for scanning subfolders and preserving their structure at the destination.

.EXAMPLE
    .\Invoke-FileNameTransformation.ps1
    Launches the application GUI from the current folder.

.EXAMPLE
    powershell.exe -NoProfile -STA -File ".\Invoke-FileNameTransformation.ps1"
    Starts the script in a single-threaded apartment (required for WPF).

.NOTES
    Requires Windows PowerShell 5.1, an STA host, and FileNameTransformation.Core.psm1 beside the script or executable.
    Profiles, logs, and the language configuration are stored
    under the current user's AppData folder, with fallback to the script directory or temporary folder when needed.
    When subfolder scanning is enabled, destination folder preservation is controlled separately and is enabled by default.
#>

[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms
Add-Type -AssemblyName Microsoft.VisualBasic

if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    Write-Warning 'Run in STA mode!'
}
# Always resolve the script's own directory fresh on every run.
# $script:AppRoot is intentionally overwritten later (line ~889) to the AppData
# folder for profiles/logs, so we keep the source location in $script:ScriptRoot.
$script:ScriptRoot = if ([string]::IsNullOrWhiteSpace($PSScriptRoot)) { (Get-Location).Path } else { $PSScriptRoot }
$script:ConfigPath = Join-Path $script:ScriptRoot 'config.json'
$script:CurrentLanguage = 'PL'
$script:Config = [pscustomobject][ordered]@{
    Version         = 2
    Language        = 'PL'
    CustomTypeRules = @()
}
$coreModulePath = Join-Path $script:ScriptRoot 'FileNameTransformation.Core.psm1'
if (-not (Test-Path -LiteralPath $coreModulePath -PathType Leaf)) {
    throw "Missing core module: $coreModulePath"
}
Import-Module $coreModulePath -Force -DisableNameChecking

if (Test-Path $script:ConfigPath) {
    try {
        Write-Verbose $script:ConfigPath
        $config = Get-Content -LiteralPath $script:ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
        $script:Config = ConvertTo-FNTConfig $config
        if ($script:Config.Language -in @('PL', 'EN', 'DE')) {
            $script:CurrentLanguage = $script:Config.Language
            Write-Information "Using saved language: $script:CurrentLanguage"
        }
    }
    catch {
        Write-Error 'Failed to load config file!'
    }
}
$customRuleValidation = Test-FNTCustomTypeRules @($script:Config.CustomTypeRules)
$script:CustomTypeRules = @($customRuleValidation.ValidRules)
foreach ($ruleError in @($customRuleValidation.Errors)) {
    Write-Warning "Custom type rule skipped: $($ruleError.Message)"
}

$script:Translations = @{
    'PL' = @{
        'WinTitle'                = 'File Name Transformer'
        'Folders_Header'          = 'Foldery'
        'Source_Folder'           = 'Folder źródłowy:'
        'Btn_Browse'              = 'Wybierz...'
        'Dest_Folder'             = 'Folder wynikowy:'
        'Lbl_Language'            = 'Język:'
        'Chk_Recursive'           = 'Skanuj także podfoldery'
        'Chk_PreserveStructure'   = 'Zachowaj strukturę folderów w folderze wynikowym'
        'Tab_Profile'             = '1. Profil'
        'Saved_Profiles'          = 'Zapisane profile'
        'Btn_New'                 = 'Nowy'
        'Btn_Load'                = 'Użyj'
        'Btn_Copy'                = 'Kopiuj'
        'Btn_Delete'              = 'Usuń'
        'Btn_Save_Config'         = 'Zapisz bieżącą konfigurację'
        'Title_Wizard'            = 'Kreator zmiany nazw plików'
        'Txt_ProfileHint'         = 'Pracuj kolejno na kartach 2–5. Wybierz istniejący profil, aby przywrócić konfigurację.'
        'Txt_ActiveProfile'       = 'Aktywny profil:'
        'Txt_Unsaved'             = '(niezapisana konfiguracja)'
        'Tab_Analysis'            = '2. Analiza plików'
        'Btn_Analyze'             = 'Analizuj folder'
        'Txt_Extension'           = 'Rozszerzenie:'
        'Header_DetectedStructs'  = 'Wykryte struktury nazw'
        'Col_Extension'           = 'Rozszerzenie'
        'Col_Structure'           = 'Struktura'
        'Col_Files'               = 'Pliki'
        'Header_Examples'         = 'Przykłady wybranego wzorca'
        'Tab_Fields'              = '3. Pola i mapowania'
        'Header_DetectedFields'   = 'Pola wykryte w nazwie'
        'Txt_FieldHint'           = 'Zaznacz pole, aby edytować jego nazwę, rolę i transformacje. Pola oznaczone V pochodzą z mapowań.'
        'Col_Sample'              = 'Przykład'
        'Col_Type'                = 'Typ'
        'Col_Name'                = 'Nazwa'
        'Col_Role'                = 'Rola'
        'Col_Source'              = 'Źródło'
        'Header_EditField'        = 'Edycja pola'
        'Txt_FieldName'           = 'Nazwa pola:'
        'Txt_Role'                = 'Rola:'
        'Role_Value'              = 'Wartość'
        'Role_Date'               = 'Data'
        'Role_Id'                 = 'Identyfikator'
        'Role_Const'              = 'Stały tekst'
        'Role_Ignore'             = 'Ignoruj'
        'Btn_ApplyField'          = 'Zastosuj zmiany pola'
        'Header_Transformations'  = 'Transformacje pola'
        'Txt_TransformHint'       = 'Transformacje modyfikują wartość pola w nazwie docelowej (np. zmiana formatu daty, przycinanie tekstu, zamiana wielkości liter).'
        'Btn_AddTransform'        = 'Dodaj transformację'
        'Btn_Remove'              = 'Usuń'
        'Header_Mappings'         = 'Mapowania (zamiana wartości z pliku CSV/TXT)'
        'Txt_MappingHint'         = 'Mapowanie zamienia wartość pola na inną wartość z pliku zewnętrznego. Tworzy nowe pole wirtualne dostępne w nazwie docelowej.'
        'Btn_AddMapping'          = 'Dodaj mapowanie'
        'Btn_Edit'                = 'Edytuj'
        'Title_EditMapping'       = 'Edytuj mapowanie'
        'Btn_Up'                  = 'W górę'
        'Btn_Down'                = 'W dół'
        'Tab_DestName'            = '4. Nazwa docelowa'
        'Header_DestElements'     = 'Elementy nazwy docelowej'
        'Btn_AddField'            = 'Dodaj pole'
        'Btn_AddText'             = 'Dodaj tekst'
        'Btn_AddSep'              = 'Dodaj _'
        'Btn_RemoveElement'       = 'Usuń element'
        'Header_Settings'         = 'Ustawienia i podgląd na żywo'
        'Chk_KeepExt'             = 'Zachowaj oryginalne rozszerzenie'
        'Txt_NewExt'              = 'Nowe rozszerzenie:'
        'Txt_PreviewTitle'        = 'Podgląd nazwy (pierwszy plik):'
        'Tab_Preview'             = '5. Podgląd i wykonanie'
        'Btn_BuildPreview'        = 'Utwórz podgląd'
        'Cbo_All'                 = 'Wszystkie'
        'Cbo_Errors'              = 'Tylko błędy'
        'Cbo_Ready'               = 'Gotowe'
        'Btn_Export'              = 'Eksportuj raport CSV'
        'Btn_OpenLog'             = 'Otwórz log'
        'Btn_Execute'             = 'Wykonaj'
        'Lbl_ExecutionMode'       = 'Akcja:'
        'Action_Copy'             = 'Kopiuj'
        'Action_Move'             = 'Przenieś'
        'Col_SourceFile'          = 'Plik źródłowy'
        'Col_DestName'            = 'Nazwa docelowa'
        'Col_Status'              = 'Status'
        'Col_Details'             = 'Szczegóły'
        'Msg_ConfirmRestart'      = 'Język został zmieniony. Uruchom aplikację ponownie, aby w pełni zastosować nowy język.'
        'Title_Info'              = 'Informacja'
        'Msg_SelectSource'        = 'Wybierz folder źródłowy'
        'Msg_SelectDest'          = 'Wybierz folder wynikowy'
        'Log_AppClosed'           = 'Aplikacja zamknięta.'
        'Status_Ready'            = 'Gotowy. Wybierz folder źródłowy i rozpocznij analizę na karcie 2.'
        'Err_Analysis'            = 'Błąd analizy'
        'Err_Field'               = 'Błąd pola'
        'Err_Transform'           = 'Błąd transformacji'
        'Err_Mapping'             = 'Błąd mapowania'
        'Err_Preview'             = 'Błąd podglądu'
        'Err_Export'              = 'Błąd eksportu'
        'Err_Execute'             = 'Błąd wykonania'
        'Err_SaveProfile'         = 'Błąd zapisu profilu'
        'Err_LoadProfile'         = 'Błąd wczytania profilu'
        'Err_CopyProfile'         = 'Błąd kopiowania profilu'
        'Err_DelProfile'          = 'Błąd usuwania profilu'
        'Title_Error'             = 'Błąd'
        'Title_Warning'           = 'Ostrzeżenie'
        'Title_Confirm'           = 'Potwierdzenie'
        'Title_SaveProfile'       = 'Zapis profilu'
        'Type_Date_1'             = 'Data (yyyyMMdd)'
        'Type_Date_2'             = 'Data (yyyy-MM-dd)'
        'Type_Date_3'             = 'Data (dd-MM-yyyy)'
        'Type_Date_4'             = 'Data (yyyy-MM)'
        'Type_Date_5'             = 'Data (MM-yyyy)'
        'Type_Date_6'             = 'Data (YYMMDD lub DDMMYY)'
        'Type_Date_7'             = 'Data (MM-DD lub DD-MM)'
        'Type_Date_8'             = 'Data (YYYY-MM-DDTHH:MM:SS lub YYYY-MM-DDTHH:MM:SSZ)'
        'Type_Num'                = 'Liczba'
        'Type_Text'               = 'Tekst'
        'Type_Integer'            = 'Liczba całkowita'
        'Type_Decimal'            = 'Liczba dziesiętna'
        'Type_DateTime'           = 'Data/czas'
        'Type_Guid'               = 'GUID'
        'Type_Version'            = 'Wersja'
        'Type_Ambiguous'          = 'Niejednoznaczny'
        'Type_Auto'               = 'Automatyczny'
        'Txt_DataType'            = 'Typ danych:'
        'Err_ResolveAmbiguous'    = 'Wybierz jednoznaczny typ danych dla pola:'
        'FieldStatus_Detected'    = 'Wykryty'
        'FieldStatus_Choice'      = 'Wymaga wyboru'
        'FieldStatus_Resolved'    = 'Rozstrzygnięty'
        'Txt_Candidates'          = 'Kandydaci:'
        'Diag_InvalidRegex'       = 'Nieprawidłowy regex tokenizera: {0}'
        'Diag_MissingSepGroup'    = "Regex tokenizera musi definiować nazwaną grupę 'sep'."
        'Diag_ZeroLength'         = 'Regex tokenizera zwrócił dopasowanie zerowej długości na pozycji {0}.'
        'Diag_IncompleteCoverage' = 'Regex tokenizera nie dopasował nazwy pliku na pozycji {0}.'
        'Diag_TokenCount'         = "Niezgodność wzorca dla '{0}': oczekiwano {1} tokenów, znaleziono {2}."
        'Diag_TokenKind'          = "Niezgodność wzorca dla '{0}', token {1}, pozycja {2}: oczekiwano {3}, znaleziono {4} '{5}'."
        'Diag_Separator'          = "Niezgodność wzorca dla '{0}', token {1}, pozycja {2}: oczekiwano separatora '{3}', znaleziono '{4}'."
        'Diag_Type'               = "Niezgodność wzorca dla '{0}', token {1}, pozycja {2}: wartość '{3}' nie ma typu '{4}'{5}."
        'Diag_KindSeparator'      = 'separatora'
        'Diag_KindValue'          = 'wartości'
        'Hint_SelectField'        = 'Wybierz pole, nadaj mu nazwę biznesową i dodaj transformacje lub mapowania.'
        'Hint_AddElements'        = 'Dodaj elementy nazwy docelowej za pomocą przycisków poniżej.'
        'Prefix_Source'           = 'Źródło'
        'Prefix_Result'           = 'Wynik'
        'Preview_Unavailable'     = 'Podgląd niedostępny'
        'Err_SrcNotExist'         = 'Wskaż istniejący folder źródłowy.'
        'Err_NoFiles'             = 'Nie znaleziono plików.'
        'Msg_Files'               = 'Pliki'
        'Msg_Structs'             = 'wykryte struktury'
        'Msg_Errors'              = 'błędy'
        'Name_Date'               = 'Data'
        'Name_Text'               = 'Tekst'
        'Name_Field'              = 'Pole'
        'Src_Name'                = 'Nazwa'
        'Src_Mapping'             = 'Mapowanie'
        'Val_Mapping'             = '(mapowanie)'
        'Title_AddTransform'      = 'Dodaj transformację'
        'Txt_TransformType'       = 'Typ transformacji:'
        'Tr_Substring'            = 'Wycinanie tekstu (Substring)'
        'Tr_DateFormat'           = 'Format daty'
        'Tr_Replace'              = 'Zamiana tekstu'
        'Tr_Case'                 = 'Zmiana wielkości liter'
        'Tr_Pad'                  = 'Dopełnienie znakami'
        'Tr_Number'               = 'Operacja matematyczna'
        'Tr_Operation'            = 'Operacja:'
        'Tr_OpAdd'                = 'Dodaj'
        'Tr_OpSubtract'           = 'Odejmij'
        'Tr_OpMultiply'           = 'Pomnóż'
        'Tr_OpDivide'             = 'Podziel'
        'Tr_OpRound'              = 'Zaokrąglij (miejsca dziesiętne)'
        'Tr_OpValue'              = 'Wartość:'
        'Tr_NumFormat'            = 'Format liczby (opcjonalnie, np. F2, D5):'
        'Tr_PosStart'             = 'Pozycja startowa (0 = początek):'
        'Tr_CharCount'            = 'Liczba znaków:'
        'Tr_FmtIn'                = 'Format wejściowy:'
        'Tr_FmtOut'               = 'Format wyjściowy:'
        'Tr_Search'               = 'Szukany tekst:'
        'Tr_NewTxt'               = 'Nowy tekst (pusty = usuń):'
        'Tr_Mode'                 = 'Tryb:'
        'Tr_Upper'                = 'WIELKIE LITERY'
        'Tr_Lower'                = 'małe litery'
        'Tr_Title'                = 'Pierwsza Wielka'
        'Tr_Side'                 = 'Strona:'
        'Tr_Left'                 = 'Z lewej'
        'Tr_Right'                = 'Z prawej'
        'Tr_PadChar'              = 'Znak dopełnienia:'
        'Tr_TargetLen'            = 'Docelowa długość:'
        'Btn_Add'                 = 'Dodaj'
        'Err_CharCount'           = 'Liczba znaków musi być > 0.'
        'Err_BothFmt'             = 'Podaj oba formaty.'
        'Err_SearchTxt'           = 'Podaj szukany tekst.'
        'Err_SelMode'             = 'Wybierz tryb.'
        'Err_SelSide'             = 'Wybierz stronę.'
        'Err_PadChar'             = 'Podaj znak dopełnienia.'
        'Err_TargetLen'           = 'Długość musi być > 0.'
        'Err_SelTransform'        = 'Wybierz typ transformacji.'
        'Err_SelOperation'        = 'Wybierz operację.'
        'Err_NumValue'            = 'Podaj prawidłową wartość liczbową.'
        'Err_DivZero'             = 'Nie można dzielić przez zero.'
        'Err_NotNumeric'          = 'Wartość pola nie jest liczbą'
        'Disp_Substr'             = 'Wycinanie: od pozycji'
        'Disp_Chars'              = 'znaków'
        'Disp_Date'               = 'Data:'
        'Disp_Replace'            = 'Zamiana:'
        'Disp_Case'               = 'Wielkość liter:'
        'Disp_Pad'                = 'Dopełnienie:'
        'Disp_Number'             = 'Matematyka:'
        'Disp_To'                 = 'do'
        'Title_AddMapping'        = 'Dodaj mapowanie'
        'Lbl_MapName'             = 'Nazwa mapowania'
        'Lbl_MapIn'               = 'Pole wejściowe'
        'Lbl_MapOut'              = 'Nazwa pola wynikowego'
        'Lbl_MapFile'             = 'Plik danych (CSV/TXT)'
        'Lbl_MapKey'              = 'Kolumna klucza'
        'Lbl_MapVal'              = 'Kolumna wartości'
        'Err_ReadHeaders'         = 'Nie udało się odczytać nagłówków:'
        'Title_FileErr'           = 'Błąd pliku'
        'Err_FillFields'          = 'Uzupełnij wszystkie pola.'
        'Title_MissingData'       = 'Brak danych'
        'Err_SelMapping'          = 'Wybierz mapowanie do edycji.'
        'Err_SelectPattern'       = 'Najpierw wybierz wzorzec na karcie 2 (Analiza plików).'
        'Err_AddDest'             = 'Dodaj elementy nazwy docelowej na karcie 4.'
        'Err_DestNotExist'        = 'Wskaż istniejący folder wynikowy.'
        'Err_MissMapFile'         = 'Brak pliku mapowania:'
        'Err_DupKey'              = 'Duplikat klucza'
        'Err_InMap'               = 'w mapowaniu'
        'Err_MissMap'             = 'Brak mapowania'
        'Err_MissField'           = 'Brak pola'
        'Err_InvalidChars'        = 'Nazwa docelowa zawiera niedozwolone znaki.'
        'Err_DupDest'             = 'Zduplikowana ścieżka docelowa'
        'Err_FileExists'          = 'Plik docelowy już istnieje'
        'Err_Blocked'             = 'Operacja zablokowana:'
        'Err_FixPrev'             = 'błędów w podglądzie. Napraw je przed kopiowaniem.'
        'Msg_CopyCount'           = 'Skopiować'
        'Msg_CopyOrig'            = 'plików? Oryginały nie zostaną zmienione.'
        'Msg_MoveOrig'            = 'plików? Oryginały zostaną przeniesione.'
        'Log_Copied'              = 'Skopiowano:'
        'Log_Moved'               = 'Przeniesiono:'
        'Log_CopyErr'             = 'Błąd kopiowania:'
        'Status_Copied'           = 'Zakończono. Skopiowano:'
        'Status_Moved'            = 'Zakończono. Przeniesiono:'
        'Lbl_ProfileName'         = 'Nazwa profilu:'
        'Status_ProfSaved'        = 'Profil zapisany:'
        'Status_ProfLoaded'       = 'Profil wczytany:'
        'Status_AnalysisDone'     = 'Analiza zakończona.'
        'Err_SelFieldTab'         = 'Zaznacz pole w tabeli.'
        'Err_FieldName'           = 'Podaj nazwę pola.'
        'Err_SelFieldTrans'       = 'Zaznacz pole, do którego chcesz dodać transformację.'
        'Err_SelFieldCombo'       = 'Wybierz pole z listy rozwijanej.'
        'Tag_Field'               = '[Pole]'
        'Tag_Text'                = '[Tekst]'
        'Tag_Separator'           = '[Separator]'
        'Status_PrevBuilt'        = 'Podgląd utworzony.'
        'Err_BuildPrev1'          = 'Najpierw utwórz podgląd.'
        'Status_Exported'         = 'Raport wyeksportowany:'
        'Err_SelProfList'         = 'Wybierz profil z listy.'
        'Err_SelProfCopy'         = 'Wybierz profil do skopiowania.'
        'Prefix_Copy'             = 'Kopia '
        'Err_SelProfDel'          = 'Wybierz profil do usunięcia.'
        'Msg_DelProf'             = 'Usunąć profil'
        'Status_ProfDel'          = 'Profil usunięty:'
        'Chk_EnforcePattern'      = 'Wymuś wzorzec na pozostałych plikach'
        'Txt_TokenRegex'          = 'Wzorzec Regex:'
        'Tip_TokenRegex'          = 'Domyślnie: (?&lt;value&gt;[^_\-\s]+)|(?&lt;sep&gt;[_\-\s]+)'
        'Tip_TokenRegexLabel'     = 'Regex musi dopasować każdy segment nazwy. Użyj (?&lt;sep&gt;...) dla separatorów; pozostałe dopasowania są wartościami. Przykład: (?&lt;value&gt;[^_\-\s]+)|(?&lt;sep&gt;[_\-\s]+).'
    }
    'EN' = @{
        'WinTitle'                = 'File Name Transformer'
        'Folders_Header'          = 'Folders'
        'Source_Folder'           = 'Source folder:'
        'Btn_Browse'              = 'Browse...'
        'Dest_Folder'             = 'Destination folder:'
        'Lbl_Language'            = 'Language:'
        'Chk_Recursive'           = 'Scan subfolders'
        'Chk_PreserveStructure'   = 'Preserve folder structure in destination folder'
        'Tab_Profile'             = '1. Profile'
        'Saved_Profiles'          = 'Saved profiles'
        'Btn_New'                 = 'New'
        'Btn_Load'                = 'Load'
        'Btn_Copy'                = 'Copy'
        'Btn_Delete'              = 'Delete'
        'Btn_Save_Config'         = 'Save current configuration'
        'Title_Wizard'            = 'File Name Transformer Wizard'
        'Txt_ProfileHint'         = 'Work sequentially on tabs 2-5. Select an existing profile to restore configuration.'
        'Txt_ActiveProfile'       = 'Active profile:'
        'Txt_Unsaved'             = '(unsaved configuration)'
        'Tab_Analysis'            = '2. File analysis'
        'Btn_Analyze'             = 'Analyze folder'
        'Txt_Extension'           = 'Extension:'
        'Header_DetectedStructs'  = 'Detected name structures'
        'Col_Extension'           = 'Extension'
        'Col_Structure'           = 'Structure'
        'Col_Files'               = 'Files'
        'Header_Examples'         = 'Examples of selected pattern'
        'Tab_Fields'              = '3. Fields and mappings'
        'Header_DetectedFields'   = 'Fields detected in name'
        'Txt_FieldHint'           = 'Select a field to edit its name, role and transformations. Fields marked with V come from mappings.'
        'Col_Sample'              = 'Sample'
        'Col_Type'                = 'Type'
        'Col_Name'                = 'Name'
        'Col_Role'                = 'Role'
        'Col_Source'              = 'Source'
        'Header_EditField'        = 'Edit field'
        'Txt_FieldName'           = 'Field name:'
        'Txt_Role'                = 'Role:'
        'Role_Value'              = 'Value'
        'Role_Date'               = 'Date'
        'Role_Id'                 = 'Identifier'
        'Role_Const'              = 'Constant text'
        'Role_Ignore'             = 'Ignore'
        'Btn_ApplyField'          = 'Apply field changes'
        'Header_Transformations'  = 'Field transformations'
        'Txt_TransformHint'       = 'Transformations modify field value in destination name (e.g., date format change, text trimming, case change).'
        'Btn_AddTransform'        = 'Add transformation'
        'Btn_Remove'              = 'Remove'
        'Header_Mappings'         = 'Mappings (value replacement from CSV/TXT)'
        'Txt_MappingHint'         = 'Mapping replaces a field value with another value from an external file. It creates a new virtual field available in destination name.'
        'Btn_AddMapping'          = 'Add mapping'
        'Btn_Edit'                = 'Edit'
        'Title_EditMapping'       = 'Edit mapping'
        'Btn_Up'                  = 'Move up'
        'Btn_Down'                = 'Move down'
        'Tab_DestName'            = '4. Destination name'
        'Header_DestElements'     = 'Destination name elements'
        'Btn_AddField'            = 'Add field'
        'Btn_AddText'             = 'Add text'
        'Btn_AddSep'              = 'Add _'
        'Btn_RemoveElement'       = 'Remove element'
        'Header_Settings'         = 'Settings and live preview'
        'Chk_KeepExt'             = 'Keep original extension'
        'Txt_NewExt'              = 'New extension:'
        'Txt_PreviewTitle'        = 'Name preview (first file):'
        'Tab_Preview'             = '5. Preview and execution'
        'Btn_BuildPreview'        = 'Build preview'
        'Cbo_All'                 = 'All'
        'Cbo_Errors'              = 'Errors only'
        'Cbo_Ready'               = 'Ready'
        'Btn_Export'              = 'Export CSV report'
        'Btn_OpenLog'             = 'Open log'
        'Btn_Execute'             = 'Execute'
        'Lbl_ExecutionMode'       = 'Action:'
        'Action_Copy'             = 'Copy'
        'Action_Move'             = 'Move'
        'Col_SourceFile'          = 'Source file'
        'Col_DestName'            = 'Destination name'
        'Col_Status'              = 'Status'
        'Col_Details'             = 'Details'
        'Msg_ConfirmRestart'      = 'Language changed. Please restart the application to fully apply the new language.'
        'Title_Info'              = 'Information'
        'Msg_SelectSource'        = 'Select source folder'
        'Msg_SelectDest'          = 'Select destination folder'
        'Log_AppClosed'           = 'Application closed.'
        'Status_Ready'            = 'Ready. Select source folder and start analysis on tab 2.'
        'Err_Analysis'            = 'Analysis error'
        'Err_Field'               = 'Field error'
        'Err_Transform'           = 'Transformation error'
        'Err_Mapping'             = 'Mapping error'
        'Err_Preview'             = 'Preview error'
        'Err_Export'              = 'Export error'
        'Err_Execute'             = 'Execution error'
        'Err_SaveProfile'         = 'Profile save error'
        'Err_LoadProfile'         = 'Profile load error'
        'Err_CopyProfile'         = 'Profile copy error'
        'Err_DelProfile'          = 'Profile delete error'
        'Title_Error'             = 'Error'
        'Title_Warning'           = 'Warning'
        'Title_Confirm'           = 'Confirmation'
        'Title_SaveProfile'       = 'Save profile'
        'Type_Date_1'             = 'Date (yyyyMMdd)'
        'Type_Date_2'             = 'Date (yyyy-MM-dd)'
        'Type_Date_3'             = 'Date (dd-MM-yyyy)'
        'Type_Date_4'             = 'Date (yyyy-MM)'
        'Type_Date_5'             = 'Date (MM-yyyy)'
        'Type_Date_6'             = 'Date (YYMMDD or DDMMYY)'
        'Type_Date_7'             = 'Date (MM-DD or DD-MM)'
        'Type_Date_8'             = 'Date (YYYY-MM-DDTHH:MM:SS or YYYY-MM-DDTHH:MM:SSZ)'
        'Type_Num'                = 'Number'
        'Type_Text'               = 'Text'
        'Type_Integer'            = 'Integer'
        'Type_Decimal'            = 'Decimal'
        'Type_DateTime'           = 'Date/time'
        'Type_Guid'               = 'GUID'
        'Type_Version'            = 'Version'
        'Type_Ambiguous'          = 'Ambiguous'
        'Type_Auto'               = 'Auto'
        'Txt_DataType'            = 'Data type:'
        'Err_ResolveAmbiguous'    = 'Select an unambiguous data type for field:'
        'FieldStatus_Detected'    = 'Detected'
        'FieldStatus_Choice'      = 'Choice required'
        'FieldStatus_Resolved'    = 'Resolved'
        'Txt_Candidates'          = 'Candidates:'
        'Diag_InvalidRegex'       = 'Invalid tokenizer regex: {0}'
        'Diag_MissingSepGroup'    = "Tokenizer regex must define the named group 'sep'."
        'Diag_ZeroLength'         = 'Tokenizer regex produced a zero-length match at position {0}.'
        'Diag_IncompleteCoverage' = 'Tokenizer regex did not match the filename at position {0}.'
        'Diag_TokenCount'         = "Pattern mismatch for '{0}': expected {1} tokens, found {2}."
        'Diag_TokenKind'          = "Pattern mismatch for '{0}', token {1}, offset {2}: expected {3}, found {4} '{5}'."
        'Diag_Separator'          = "Pattern mismatch for '{0}', token {1}, offset {2}: expected separator '{3}', found '{4}'."
        'Diag_Type'               = "Pattern mismatch for '{0}', token {1}, offset {2}: value '{3}' is not type '{4}'{5}."
        'Diag_KindSeparator'      = 'separator'
        'Diag_KindValue'          = 'value'
        'Hint_SelectField'        = 'Select a field, give it a business name, and add transformations or mappings.'
        'Hint_AddElements'        = 'Add destination name elements using buttons below.'
        'Prefix_Source'           = 'Source'
        'Prefix_Result'           = 'Result'
        'Preview_Unavailable'     = 'Preview unavailable'
        'Err_SrcNotExist'         = 'Specify an existing source folder.'
        'Err_NoFiles'             = 'No files found.'
        'Msg_Files'               = 'Files'
        'Msg_Structs'             = 'detected structures'
        'Msg_Errors'              = 'errors'
        'Name_Date'               = 'Date'
        'Name_Text'               = 'Text'
        'Name_Field'              = 'Field'
        'Src_Name'                = 'Name'
        'Src_Mapping'             = 'Mapping'
        'Val_Mapping'             = '(mapping)'
        'Title_AddTransform'      = 'Add transformation'
        'Txt_TransformType'       = 'Transformation type:'
        'Tr_Substring'            = 'Text trimming (Substring)'
        'Tr_DateFormat'           = 'Date format'
        'Tr_Replace'              = 'Text replacement'
        'Tr_Case'                 = 'Case change'
        'Tr_Pad'                  = 'Padding'
        'Tr_Number'               = 'Math operation'
        'Tr_Operation'            = 'Operation:'
        'Tr_OpAdd'                = 'Add'
        'Tr_OpSubtract'           = 'Subtract'
        'Tr_OpMultiply'           = 'Multiply'
        'Tr_OpDivide'             = 'Divide'
        'Tr_OpRound'              = 'Round (decimal places)'
        'Tr_OpValue'              = 'Value:'
        'Tr_NumFormat'            = 'Number format (optional, e.g. F2, D5):'
        'Tr_PosStart'             = 'Start position (0 = start):'
        'Tr_CharCount'            = 'Number of characters:'
        'Tr_FmtIn'                = 'Input format:'
        'Tr_FmtOut'               = 'Output format:'
        'Tr_Search'               = 'Search text:'
        'Tr_NewTxt'               = 'New text (empty = remove):'
        'Tr_Mode'                 = 'Mode:'
        'Tr_Upper'                = 'UPPERCASE'
        'Tr_Lower'                = 'lowercase'
        'Tr_Title'                = 'Title Case'
        'Tr_Side'                 = 'Side:'
        'Tr_Left'                 = 'Left'
        'Tr_Right'                = 'Right'
        'Tr_PadChar'              = 'Pad character:'
        'Tr_TargetLen'            = 'Target length:'
        'Btn_Add'                 = 'Add'
        'Err_CharCount'           = 'Number of characters must be > 0.'
        'Err_BothFmt'             = 'Provide both formats.'
        'Err_SearchTxt'           = 'Provide search text.'
        'Err_SelMode'             = 'Select mode.'
        'Err_SelSide'             = 'Select side.'
        'Err_PadChar'             = 'Provide pad character.'
        'Err_TargetLen'           = 'Length must be > 0.'
        'Err_SelTransform'        = 'Select transformation type.'
        'Err_SelOperation'        = 'Select an operation.'
        'Err_NumValue'            = 'Provide a valid numeric value.'
        'Err_DivZero'             = 'Cannot divide by zero.'
        'Err_NotNumeric'          = 'Field value is not numeric'
        'Disp_Substr'             = 'Trimming: from pos'
        'Disp_Chars'              = 'chars'
        'Disp_Date'               = 'Date:'
        'Disp_Replace'            = 'Replace:'
        'Disp_Case'               = 'Case:'
        'Disp_Pad'                = 'Pad:'
        'Disp_Number'             = 'Math:'
        'Disp_To'                 = 'to'
        'Title_AddMapping'        = 'Add mapping'
        'Lbl_MapName'             = 'Mapping name'
        'Lbl_MapIn'               = 'Input field'
        'Lbl_MapOut'              = 'Output field name'
        'Lbl_MapFile'             = 'Data file (CSV/TXT)'
        'Lbl_MapKey'              = 'Key column'
        'Lbl_MapVal'              = 'Value column'
        'Err_ReadHeaders'         = 'Failed to read headers:'
        'Title_FileErr'           = 'File error'
        'Err_FillFields'          = 'Fill in all fields.'
        'Title_MissingData'       = 'Missing data'
        'Err_SelectPattern'       = 'First select a pattern on tab 2 (File analysis).'
        'Err_AddDest'             = 'Add destination name elements on tab 4.'
        'Err_DestNotExist'        = 'Specify an existing destination folder.'
        'Err_MissMapFile'         = 'Missing mapping file:'
        'Err_DupKey'              = 'Duplicate key'
        'Err_InMap'               = 'in mapping'
        'Err_MissMap'             = 'Missing mapping'
        'Err_MissField'           = 'Missing field'
        'Err_InvalidChars'        = 'Destination name contains invalid characters.'
        'Err_DupDest'             = 'Duplicate destination path'
        'Err_FileExists'          = 'Destination file already exists'
        'Err_Blocked'             = 'Operation blocked:'
        'Err_FixPrev'             = 'errors in preview. Fix them before copying.'
        'Msg_CopyCount'           = 'Copy'
        'Msg_CopyOrig'            = 'files? Originals will not be modified.'
        'Log_Copied'              = 'Copied:'
        'Log_CopyErr'             = 'Copy error:'
        'Status_Copied'           = 'Finished. Copied:'
        'Lbl_ProfileName'         = 'Profile name:'
        'Status_ProfSaved'        = 'Profile saved:'
        'Status_ProfLoaded'       = 'Profile loaded:'
        'Status_AnalysisDone'     = 'Analysis complete.'
        'Err_SelFieldTab'         = 'Select a field in the table.'
        'Err_FieldName'           = 'Provide field name.'
        'Err_SelFieldTrans'       = 'Select a field to add transformation to.'
        'Err_SelFieldCombo'       = 'Select a field from dropdown.'
        'Tag_Field'               = '[Field]'
        'Tag_Text'                = '[Text]'
        'Tag_Separator'           = '[Separator]'
        'Status_PrevBuilt'        = 'Preview built.'
        'Err_BuildPrev1'          = 'Build preview first.'
        'Status_Exported'         = 'Report exported:'
        'Err_SelProfList'         = 'Select a profile from list.'
        'Err_SelProfCopy'         = 'Select a profile to copy.'
        'Prefix_Copy'             = 'Copy of '
        'Err_SelProfDel'          = 'Select a profile to delete.'
        'Msg_DelProf'             = 'Delete profile'
        'Status_ProfDel'          = 'Profile deleted:'
        'Chk_EnforcePattern'      = 'Enforce pattern on other files'
        'Txt_TokenRegex'          = 'Regex Pattern:'
        'Tip_TokenRegex'          = 'Default: (?&lt;value&gt;[^_\-\s]+)|(?&lt;sep&gt;[_\-\s]+)'
        'Tip_TokenRegexLabel'     = 'The regex must match every filename segment. Use (?&lt;sep&gt;...) for separators; all other matches are values. Example: (?&lt;value&gt;[^_\-\s]+)|(?&lt;sep&gt;[_\-\s]+).'
    }
    'DE' = @{
        'WinTitle'                = 'File Name Transformer'
        'Folders_Header'          = 'Ordner'
        'Source_Folder'           = 'Quellordner:'
        'Btn_Browse'              = 'Durchsuchen...'
        'Dest_Folder'             = 'Zielordner:'
        'Lbl_Language'            = 'Sprache:'
        'Chk_Recursive'           = 'Unterordner scannen'
        'Chk_PreserveStructure'   = 'Ordnerstruktur im Zielordner beibehalten'
        'Tab_Profile'             = '1. Profil'
        'Saved_Profiles'          = 'Gespeicherte Profile'
        'Btn_New'                 = 'Neu'
        'Btn_Load'                = 'Laden'
        'Btn_Copy'                = 'Kopieren'
        'Btn_Delete'              = 'Löschen'
        'Btn_Save_Config'         = 'Aktuelle Konfiguration speichern'
        'Title_Wizard'            = 'Dateinamen-Transformator-Assistent'
        'Txt_ProfileHint'         = 'Arbeiten Sie nacheinander an den Registerkarten 2-5. Wählen Sie ein vorhandenes Profil aus, um die Konfiguration wiederherzustellen.'
        'Txt_ActiveProfile'       = 'Aktives Profil:'
        'Txt_Unsaved'             = '(ungespeicherte Konfiguration)'
        'Tab_Analysis'            = '2. Dateianalyse'
        'Btn_Analyze'             = 'Ordner analysieren'
        'Txt_Extension'           = 'Erweiterung:'
        'Header_DetectedStructs'  = 'Erkannte Namensstrukturen'
        'Col_Extension'           = 'Erw.'
        'Col_Structure'           = 'Struktur'
        'Col_Files'               = 'Dateien'
        'Header_Examples'         = 'Beispiele für das ausgewählte Muster'
        'Tab_Fields'              = '3. Felder und Zuordnungen'
        'Header_DetectedFields'   = 'Im Namen erkannte Felder'
        'Txt_FieldHint'           = 'Wählen Sie ein Feld aus, um seinen Namen, seine Rolle und seine Transformationen zu bearbeiten. Mit V markierte Felder stammen aus Zuordnungen.'
        'Col_Sample'              = 'Beispiel'
        'Col_Type'                = 'Typ'
        'Col_Name'                = 'Name'
        'Col_Role'                = 'Rolle'
        'Col_Source'              = 'Quelle'
        'Header_EditField'        = 'Feld bearbeiten'
        'Txt_FieldName'           = 'Feldname:'
        'Txt_Role'                = 'Rolle:'
        'Role_Value'              = 'Wert'
        'Role_Date'               = 'Datum'
        'Role_Id'                 = 'Identifikator'
        'Role_Const'              = 'Konstanter Text'
        'Role_Ignore'             = 'Ignorieren'
        'Btn_ApplyField'          = 'Feldänderungen anwenden'
        'Header_Transformations'  = 'Feldtransformationen'
        'Txt_TransformHint'       = 'Transformationen ändern den Feldwert im Zielnamen (z. B. Änderung des Datumsformats, Textkürzung, Groß-/Kleinschreibung).'
        'Btn_AddTransform'        = 'Transformation hinzufügen'
        'Btn_Remove'              = 'Entfernen'
        'Header_Mappings'         = 'Zuordnungen (Wertersatz aus CSV/TXT)'
        'Txt_MappingHint'         = 'Die Zuordnung ersetzt einen Feldwert durch einen anderen Wert aus einer externen Datei. Es wird ein neues virtuelles Feld erstellt.'
        'Btn_AddMapping'          = 'Zuordnung hinzufügen'
        'Btn_Edit'                = 'Bearbeiten'
        'Title_EditMapping'       = 'Zuordnung bearbeiten'
        'Btn_Up'                  = 'Nach oben'
        'Btn_Down'                = 'Nach unten'
        'Tab_DestName'            = '4. Zielname'
        'Header_DestElements'     = 'Elemente des Zielnamens'
        'Btn_AddField'            = 'Feld hinzufügen'
        'Btn_AddText'             = 'Text hinzufügen'
        'Btn_AddSep'              = '_ hinzufügen'
        'Btn_RemoveElement'       = 'Element entfernen'
        'Header_Settings'         = 'Einstellungen und Live-Vorschau'
        'Chk_KeepExt'             = 'Original-Erweiterung beibehalten'
        'Txt_NewExt'              = 'Neue Erweiterung:'
        'Txt_PreviewTitle'        = 'Namensvorschau (erste Datei):'
        'Tab_Preview'             = '5. Vorschau und Ausführung'
        'Btn_BuildPreview'        = 'Vorschau erstellen'
        'Cbo_All'                 = 'Alle'
        'Cbo_Errors'              = 'Nur Fehler'
        'Cbo_Ready'               = 'Bereit'
        'Btn_Export'              = 'CSV-Bericht exportieren'
        'Btn_OpenLog'             = 'Protokoll öffnen'
        'Btn_Execute'             = 'Ausführen'
        'Lbl_ExecutionMode'       = 'Aktion:'
        'Action_Copy'             = 'Kopieren'
        'Action_Move'             = 'Verschieben'
        'Col_SourceFile'          = 'Quelldatei'
        'Col_DestName'            = 'Zielname'
        'Col_Status'              = 'Status'
        'Col_Details'             = 'Details'
        'Msg_ConfirmRestart'      = 'Sprache geändert. Bitte starten Sie die Anwendung neu, um die neue Sprache vollständig anzuwenden.'
        'Title_Info'              = 'Information'
        'Msg_SelectSource'        = 'Quellordner auswählen'
        'Msg_SelectDest'          = 'Zielordner auswählen'
        'Log_AppClosed'           = 'Anwendung geschlossen.'
        'Status_Ready'            = 'Bereit. Quellordner auswählen und Analyse auf Registerkarte 2 starten.'
        'Err_Analysis'            = 'Analysefehler'
        'Err_Field'               = 'Feldfehler'
        'Err_Transform'           = 'Transformationsfehler'
        'Err_Mapping'             = 'Zuordnungsfehler'
        'Err_Preview'             = 'Vorschaufehler'
        'Err_Export'              = 'Exportfehler'
        'Err_Execute'             = 'Ausführungsfehler'
        'Err_SaveProfile'         = 'Profil-Speicherfehler'
        'Err_LoadProfile'         = 'Profil-Ladefehler'
        'Err_CopyProfile'         = 'Profil-Kopierfehler'
        'Err_DelProf'             = 'Profil-Löschfehler'
        'Title_Error'             = 'Fehler'
        'Title_Warning'           = 'Warnung'
        'Title_Confirm'           = 'Bestätigung'
        'Title_SaveProfile'       = 'Profil speichern'
        'Type_Date_1'             = 'Datum (yyyyMMdd)'
        'Type_Date_2'             = 'Datum (yyyy-MM-dd)'
        'Type_Date_3'             = 'Datum (dd-MM-yyyy)'
        'Type_Date_4'             = 'Datum (yyyy-MM)'
        'Type_Date_5'             = 'Datum (MM-yyyy)'
        'Type_Date_6'             = 'Datum (YYMMDD oder DDMMYY)'
        'Type_Date_7'             = 'Datum (MM-DD oder DD-MM)'
        'Type_Date_8'             = 'Datum (YYYY-MM-DDTHH:MM:SS oder YYYY-MM-DDTHH:MM:SSZ)'
        'Type_Num'                = 'Zahl'
        'Type_Text'               = 'Text'
        'Type_Integer'            = 'Ganzzahl'
        'Type_Decimal'            = 'Dezimalzahl'
        'Type_DateTime'           = 'Datum/Uhrzeit'
        'Type_Guid'               = 'GUID'
        'Type_Version'            = 'Version'
        'Type_Ambiguous'          = 'Mehrdeutig'
        'Type_Auto'               = 'Automatisch'
        'Txt_DataType'            = 'Datentyp:'
        'Err_ResolveAmbiguous'    = 'Wählen Sie einen eindeutigen Datentyp für das Feld:'
        'FieldStatus_Detected'    = 'Erkannt'
        'FieldStatus_Choice'      = 'Auswahl erforderlich'
        'FieldStatus_Resolved'    = 'Aufgelöst'
        'Txt_Candidates'          = 'Kandidaten:'
        'Diag_InvalidRegex'       = 'Ungültiger Tokenizer-Regex: {0}'
        'Diag_MissingSepGroup'    = "Der Tokenizer-Regex muss die benannte Gruppe 'sep' definieren."
        'Diag_ZeroLength'         = 'Der Tokenizer-Regex erzeugte an Position {0} einen Treffer der Länge null.'
        'Diag_IncompleteCoverage' = 'Der Tokenizer-Regex konnte den Dateinamen an Position {0} nicht zuordnen.'
        'Diag_TokenCount'         = "Musterabweichung für '{0}': {1} Token erwartet, {2} gefunden."
        'Diag_TokenKind'          = "Musterabweichung für '{0}', Token {1}, Position {2}: {3} erwartet, {4} '{5}' gefunden."
        'Diag_Separator'          = "Musterabweichung für '{0}', Token {1}, Position {2}: Trennzeichen '{3}' erwartet, '{4}' gefunden."
        'Diag_Type'               = "Musterabweichung für '{0}', Token {1}, Position {2}: Wert '{3}' entspricht nicht dem Typ '{4}'{5}."
        'Diag_KindSeparator'      = 'Trennzeichen'
        'Diag_KindValue'          = 'Wert'
        'Hint_SelectField'        = 'Wählen Sie ein Feld aus, geben Sie ihm einen Geschäftsnamen und fügen Sie Transformationen oder Zuordnungen hinzu.'
        'Hint_AddElements'        = 'Fügen Sie Zielnamenselemente mit den Schaltflächen unten hinzu.'
        'Prefix_Source'           = 'Quelle'
        'Prefix_Result'           = 'Ergebnis'
        'Preview_Unavailable'     = 'Vorschau nicht verfügbar'
        'Err_SrcNotExist'         = 'Geben Sie einen vorhandenen Quellordner an.'
        'Err_NoFiles'             = 'Keine Dateien gefunden.'
        'Msg_Files'               = 'Dateien'
        'Msg_Structs'             = 'erkannte Strukturen'
        'Msg_Errors'              = 'Fehler'
        'Name_Date'               = 'Datum'
        'Name_Text'               = 'Text'
        'Name_Field'              = 'Feld'
        'Src_Name'                = 'Name'
        'Src_Mapping'             = 'Zuordnung'
        'Val_Mapping'             = '(Zuordnung)'
        'Title_AddTransform'      = 'Transformation hinzufügen'
        'Txt_TransformType'       = 'Transformationstyp:'
        'Tr_Substring'            = 'Textkürzung (Substring)'
        'Tr_DateFormat'           = 'Datumsformat'
        'Tr_Replace'              = 'Textersatz'
        'Tr_Case'                 = 'Groß-/Kleinschreibung'
        'Tr_Pad'                  = 'Auffüllen'
        'Tr_Number'               = 'Rechenoperation'
        'Tr_Operation'            = 'Operation:'
        'Tr_OpAdd'                = 'Addieren'
        'Tr_OpSubtract'           = 'Subtrahieren'
        'Tr_OpMultiply'           = 'Multiplizieren'
        'Tr_OpDivide'             = 'Dividieren'
        'Tr_OpRound'              = 'Runden (Dezimalstellen)'
        'Tr_OpValue'              = 'Wert:'
        'Tr_NumFormat'            = 'Zahlenformat (optional, z. B. F2, D5):'
        'Tr_PosStart'             = 'Startposition (0 = Anfang):'
        'Tr_CharCount'            = 'Anzahl der Zeichen:'
        'Tr_FmtIn'                = 'Eingabeformat:'
        'Tr_FmtOut'               = 'Ausgabeformat:'
        'Tr_Search'               = 'Suchtext:'
        'Tr_NewTxt'               = 'Neuer Text (leer = entfernen):'
        'Tr_Mode'                 = 'Modus:'
        'Tr_Upper'                = 'GROSSBUCHSTABEN'
        'Tr_Lower'                = 'kleinbuchstaben'
        'Tr_Title'                = 'Erster Buchstabe groß'
        'Tr_Side'                 = 'Seite:'
        'Tr_Left'                 = 'Links'
        'Tr_Right'                = 'Rechts'
        'Tr_PadChar'              = 'Füllzeichen:'
        'Tr_TargetLen'            = 'Ziellänge:'
        'Btn_Add'                 = 'Hinzufügen'
        'Err_CharCount'           = 'Die Anzahl der Zeichen muss > 0 sein.'
        'Err_BothFmt'             = 'Geben Sie beide Formate an.'
        'Err_SearchTxt'           = 'Geben Sie den Suchtext an.'
        'Err_SelMode'             = 'Wählen Sie den Modus.'
        'Err_SelSide'             = 'Wählen Sie die Seite.'
        'Err_PadChar'             = 'Geben Sie das Füllzeichen an.'
        'Err_TargetLen'           = 'Länge muss > 0 sein.'
        'Err_SelTransform'        = 'Wählen Sie den Transformationstyp.'
        'Err_SelOperation'        = 'Wählen Sie eine Operation.'
        'Err_NumValue'            = 'Geben Sie einen gültigen numerischen Wert an.'
        'Err_DivZero'             = 'Division durch Null nicht möglich.'
        'Err_NotNumeric'          = 'Feldwert ist keine Zahl'
        'Disp_Substr'             = 'Kürzen: ab Pos'
        'Disp_Chars'              = 'Zeichen'
        'Disp_Date'               = 'Datum:'
        'Disp_Replace'            = 'Ersatz:'
        'Disp_Case'               = 'Gr./Kl.:'
        'Disp_Pad'                = 'Auffüllen:'
        'Disp_Number'             = 'Rechnen:'
        'Disp_To'                 = 'bis'
        'Title_AddMapping'        = 'Zuordnung hinzufügen'
        'Lbl_MapName'             = 'Zuordnungsname'
        'Lbl_MapIn'               = 'Eingabefeld'
        'Lbl_MapOut'              = 'Ausgabefeldname'
        'Lbl_MapFile'             = 'Datendatei (CSV/TXT)'
        'Lbl_MapKey'              = 'Schlüsselspalte'
        'Lbl_MapVal'              = 'Wertspalte'
        'Err_ReadHeaders'         = 'Fehler beim Lesen der Kopfzeilen:'
        'Title_FileErr'           = 'Dateifehler'
        'Err_FillFields'          = 'Füllen Sie alle Felder aus.'
        'Title_MissingData'       = 'Fehlende Daten'
        'Err_SelectPattern'       = 'Wählen Sie zunächst ein Muster auf Registerkarte 2 (Dateianalyse) aus.'
        'Err_AddDest'             = 'Fügen Sie Zielnamenselemente auf Registerkarte 4 hinzu.'
        'Err_DestNotExist'        = 'Geben Sie einen vorhandenen Zielordner an.'
        'Err_MissMapFile'         = 'Fehlende Zuordnungsdatei:'
        'Err_DupKey'              = 'Doppelter Schlüssel'
        'Err_InMap'               = 'in Zuordnung'
        'Err_MissMap'             = 'Fehlende Zuordnung'
        'Err_MissField'           = 'Fehlendes Feld'
        'Err_InvalidChars'        = 'Der Zielname enthält ungültige Zeichen.'
        'Err_DupDest'             = 'Doppelter Zielpfad'
        'Err_FileExists'          = 'Zieldatei existiert bereits'
        'Err_Blocked'             = 'Vorgang blockiert:'
        'Err_FixPrev'             = 'Fehler in der Vorschau. Beheben Sie diese vor dem Kopieren.'
        'Msg_CopyCount'           = 'Kopieren'
        'Msg_CopyOrig'            = 'Dateien? Originale werden nicht geändert.'
        'Msg_MoveOrig'            = 'Dateien? Originale werden verschoben.'
        'Log_Copied'              = 'Kopiert:'
        'Log_Moved'               = 'Verschoben:'
        'Log_CopyErr'             = 'Kopierfehler:'
        'Status_Copied'           = 'Abgeschlossen. Kopiert:'
        'Status_Moved'            = 'Abgeschlossen. Verschoben:'
        'Lbl_ProfileName'         = 'Profilname:'
        'Status_ProfSaved'        = 'Profil gespeichert:'
        'Status_ProfLoaded'       = 'Profil geladen:'
        'Status_AnalysisDone'     = 'Analyse abgeschlossen.'
        'Err_SelFieldTab'         = 'Wählen Sie ein Feld in der Tabelle.'
        'Err_FieldName'           = 'Feldname angeben.'
        'Err_SelFieldTrans'       = 'Wählen Sie ein Feld aus, dem Sie eine Transformation hinzufügen möchten.'
        'Err_SelFieldCombo'       = 'Wählen Sie ein Feld aus dem Dropdown-Menü.'
        'Tag_Field'               = '[Feld]'
        'Tag_Text'                = '[Text]'
        'Tag_Separator'           = '[Trennzeichen]'
        'Status_PrevBuilt'        = 'Vorschau erstellt.'
        'Err_BuildPrev1'          = 'Zuerst Vorschau erstellen.'
        'Status_Exported'         = 'Bericht exportiert:'
        'Err_SelProfList'         = 'Wählen Sie ein Profil aus der Liste.'
        'Err_SelProfCopy'         = 'Wählen Sie ein Profil zum Kopieren aus.'
        'Prefix_Copy'             = 'Kopie von '
        'Err_SelProfDel'          = 'Wählen Sie ein Profil zum Löschen aus.'
        'Msg_DelProf'             = 'Profil löschen'
        'Status_ProfDel'          = 'Profil gelöscht:'
        'Chk_EnforcePattern'      = 'Muster auf andere Dateien erzwingen'
        'Txt_TokenRegex'          = 'Regex-Muster:'
        'Tip_TokenRegex'          = 'Standard: (?&lt;value&gt;[^_\-\s]+)|(?&lt;sep&gt;[_\-\s]+)'
        'Tip_TokenRegexLabel'     = 'Der Regex muss jedes Segment des Dateinamens erfassen. Verwenden Sie (?&lt;sep&gt;...) für Trennzeichen; alle anderen Treffer sind Werte. Beispiel: (?&lt;value&gt;[^_\-\s]+)|(?&lt;sep&gt;[_\-\s]+).'
    }
}
# Load translations
. (Join-Path $script:ScriptRoot 'src\Translations.ps1')

# Set application language based on OS or profile (fallback to EN)
$osLang = (Get-Culture).TwoLetterISOLanguageName.ToUpper()
$script:CurrentLanguage = if ($osLang -in @('PL', 'EN', 'DE')) { $osLang } else { 'EN' }

function T([string]$Key) {
    if ($script:Translations[$script:CurrentLanguage].ContainsKey($Key)) {
        return $script:Translations[$script:CurrentLanguage][$Key]
    }
    return $Key
}
#endregion
#region Application Directories and State
$baseAppRoot = if ($env:APPDATA) { Join-Path $env:APPDATA 'FileNameTransformer' } elseif ($script:AppRoot) { $script:AppRoot } else { (Get-Location).Path }
if ([string]::IsNullOrWhiteSpace($baseAppRoot)) {
    $baseAppRoot = [System.IO.Path]::GetTempPath()
}
$script:AppRoot = $baseAppRoot
$script:ProfileRoot = Join-Path $script:AppRoot 'Profiles'
$script:LogRoot = Join-Path $script:AppRoot 'Logs'
New-Item -ItemType Directory -Path $script:ProfileRoot, $script:LogRoot -Force | Out-Null
$script:LogPath = Join-Path $script:LogRoot ('FileNameTransformer_{0:yyyyMMdd_HHmmss}.log' -f (Get-Date))

$script:Patterns = @()
$script:Fields = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
$script:Mappings = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
$script:OutputParts = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
$script:PreviewRows = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
$script:CurrentProfileName = ''
$script:CurrentPattern = $null
#endregion

#region Screen Size
$script:ScreenWidth = [System.Windows.SystemParameters]::PrimaryScreenWidth * 0.9
$script:ScreenHeight = [System.Windows.SystemParameters]::PrimaryScreenHeight * 0.9
#endregion

#region XAML Definition
# Read MainWindow.xaml from disk
$xamlPath = Join-Path $script:ScriptRoot 'MainWindow.xaml'
if (-not (Test-Path -LiteralPath $xamlPath -PathType Leaf)) {
    throw "Missing UI template: $xamlPath"
}
$xamlTemplate = [System.IO.File]::ReadAllText($xamlPath)
#endregion

#region Window Creation and Element Binding
$xamlTemplate = $xamlTemplate.Replace('{ScreenWidth}', [string][int]$script:ScreenWidth)
$xamlTemplate = $xamlTemplate.Replace('{ScreenHeight}', [string][int]$script:ScreenHeight)
foreach ($key in $script:Translations[$script:CurrentLanguage].Keys) {
    $xamlTemplate = $xamlTemplate.Replace("{t:$key}", $script:Translations[$script:CurrentLanguage][$key])
}
[xml]$xamlDoc = $xamlTemplate
$reader = New-Object System.Xml.XmlNodeReader $xamlDoc
$window = [Windows.Markup.XamlReader]::Load($reader)

# Bind all named XAML elements to script-scoped variables
$ns = New-Object System.Xml.XmlNamespaceManager $xamlDoc.NameTable
$ns.AddNamespace('x', 'http://schemas.microsoft.com/winfx/2006/xaml')
$xamlDoc.SelectNodes('//*[@x:Name]', $ns) | ForEach-Object {
    $name = $_.GetAttribute('Name', 'http://schemas.microsoft.com/winfx/2006/xaml')
    Set-Variable -Name $name -Value $window.FindName($name) -Scope Script
}
#endregion

#region Utility Functions
# Load source modules
. (Join-Path $script:ScriptRoot 'src\UI.ps1')
. (Join-Path $script:ScriptRoot 'src\Analysis.ps1')
. (Join-Path $script:ScriptRoot 'src\Transforms.ps1')
. (Join-Path $script:ScriptRoot 'src\Mappings.ps1')
. (Join-Path $script:ScriptRoot 'src\Preview.ps1')
. (Join-Path $script:ScriptRoot 'src\Profiles.ps1')

function Log([string]$message, [string]$level = 'INFO') {
    if ([string]::IsNullOrWhiteSpace($script:LogPath)) {
        if ([string]::IsNullOrWhiteSpace($script:LogRoot)) {
            $script:LogRoot = Join-Path $script:AppRoot 'Logs'
        }
        if ([string]::IsNullOrWhiteSpace($script:AppRoot)) {
            $script:AppRoot = [System.IO.Path]::GetTempPath()
        }
        New-Item -ItemType Directory -Path $script:LogRoot -Force | Out-Null
        $script:LogPath = Join-Path $script:LogRoot ('FileNameTransformer_{0:yyyyMMdd_HHmmss}.log' -f (Get-Date))
    }

    $line = '{0:yyyy-MM-dd HH:mm:ss} [{1}] {2}' -f (Get-Date), $level, $message
    Add-Content -LiteralPath $script:LogPath -Encoding UTF8 -Value $line
}

function SetStatus([string]$message) {
    $StatusText.Text = $message
    Log $message
}

function GetLocalizedErrorMessage([Exception]$exception) {
    if ($null -eq $exception -or -not $exception.Data.Contains('FNTCode')) {
        return $exception.Message
    }

    $kindLabel = {
        param($kind)
        if ($kind -eq 'separator') { return (T 'Diag_KindSeparator') }
        return (T 'Diag_KindValue')
    }
    switch ([string]$exception.Data['FNTCode']) {
        'Tokenizer.InvalidRegex' {
            return ((T 'Diag_InvalidRegex') -f $exception.Data['Reason'])
        }
        'Tokenizer.MissingSeparatorGroup' { return (T 'Diag_MissingSepGroup') }
        'Tokenizer.ZeroLength' {
            return ((T 'Diag_ZeroLength') -f $exception.Data['Position'])
        }
        'Tokenizer.IncompleteCoverage' {
            return ((T 'Diag_IncompleteCoverage') -f $exception.Data['Position'])
        }
        'Pattern.TokenCount' {
            return ((T 'Diag_TokenCount') -f $exception.Data['Name'], $exception.Data['Expected'], $exception.Data['Actual'])
        }
        'Pattern.TokenKind' {
            return ((T 'Diag_TokenKind') -f $exception.Data['Name'], $exception.Data['Token'], $exception.Data['Offset'],
                (& $kindLabel $exception.Data['Expected']), (& $kindLabel $exception.Data['Actual']), $exception.Data['Value'])
        }
        'Pattern.Separator' {
            return ((T 'Diag_Separator') -f $exception.Data['Name'], $exception.Data['Token'], $exception.Data['Offset'],
                $exception.Data['Expected'], $exception.Data['Actual'])
        }
        'Pattern.Type' {
            $formatText = if ($exception.Data['Format']) { " ($($exception.Data['Format']))" } else { '' }
            return ((T 'Diag_Type') -f $exception.Data['Name'], $exception.Data['Token'], $exception.Data['Offset'],
                $exception.Data['Value'], $exception.Data['TypeId'], $formatText)
        }
        default { return $exception.Message }
    }
}

function ErrorBox([string]$title, $err) {
    $localizedMessage = GetLocalizedErrorMessage $err.Exception
    $msg = "Komunikat: $localizedMessage"
    $msg += "`nLinia:     $($err.InvocationInfo.ScriptLineNumber)"
    $msg += "`nLog:       $script:LogPath"
    Log "$title | $localizedMessage" 'ERROR'
    [Windows.MessageBox]::Show($msg, $title, 'OK', 'Error') | Out-Null
}

function FolderDialog([string]$caption, [string]$initial) {
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = $caption

    $initialPath = [string]$initial
    if (-not [string]::IsNullOrWhiteSpace($initialPath) -and (Test-Path -LiteralPath $initialPath -PathType Container)) {
        $dlg.SelectedPath = $initialPath
    }

    if ($dlg.ShowDialog() -eq 'OK') { $dlg.SelectedPath }
}

function FileDialog {
    $dlg = New-Object Microsoft.Win32.OpenFileDialog
    $dlg.Filter = 'CSV (*.csv)|*.csv|Tekst (*.txt)|*.txt|Wszystkie (*.*)|*.*'
    if ($dlg.ShowDialog()) { $dlg.FileName }
}

function UpdateUI {
    # Force WPF dispatcher to process pending UI updates
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
        [Action] {},
        [System.Windows.Threading.DispatcherPriority]::Background
    )
}

#endregion

#region Event Handlers
# Load event handlers
. (Join-Path $script:ScriptRoot 'src\Events.ps1')

# --- Folder browse ---
$BrowseSource.Add_Click({
        $p = FolderDialog (T 'Msg_SelectSource') $SourcePath.Text
        if ($p) { $SourcePath.Text = $p }
    })
$BrowseDestination.Add_Click({
        $p = FolderDialog (T 'Msg_SelectDest') $DestinationPath.Text
        if ($p) { $DestinationPath.Text = $p }
    })

# --- Tab 2: Analysis ---
$Analyze.Add_Click({
        try { AnalyzePatterns; SetStatus (T 'Status_AnalysisDone') }
        catch { ErrorBox (T 'Err_Analysis') $_ }
    })
$ExtensionFilter.Add_SelectionChanged({
        try { BuildPatternList } catch {}
    })
$PatternGrid.Add_SelectionChanged({
        if ($PatternGrid.SelectedItem) {
            SetPattern $PatternGrid.SelectedItem
        }
    })

# --- Tab 3: Fields ---
$FieldGrid.Add_SelectionChanged({
        $f = $FieldGrid.SelectedItem
        if ($f) {
            $FieldName.Text = $f.Name
            # Select matching role item
            $FieldRole.SelectedItem = @($FieldRole.Items | Where-Object { $_.Content -eq $f.Role })[0]
            if (-not $FieldRole.SelectedItem) { $FieldRole.SelectedIndex = 0 }
            $typeOptions = @(GetFieldTypeOptions $f)
            $FieldType.ItemsSource = $typeOptions
            $selectedType = @($typeOptions | Where-Object {
                    $_.Id -eq $f.SelectedTypeId -and [string]$_.Format -eq [string]$f.SelectedFormat
                } | Select-Object -First 1)
            if ($selectedType.Count -gt 0) { $FieldType.SelectedItem = $selectedType[0] }
            else { $FieldType.SelectedIndex = 0 }
            $CandidateInfo.Text = GetFieldCandidateSummary $f
            # Show transforms for this field
            $TransformList.ItemsSource = $null
            if ($f.Transforms) {
                $TransformList.ItemsSource = @($f.Transforms)
            }
        }
    })

$FieldApply.Add_Click({
        try {
            $f = $FieldGrid.SelectedItem
            if (-not $f) { throw (T 'Err_SelFieldTab') }
            $newName = $FieldName.Text.Trim()
            if (-not $newName) { throw (T 'Err_FieldName') }
            $oldName = $f.Name

            $f.Name = $newName
            $f.Role = [string]$FieldRole.SelectedItem.Content
            $selectedType = $FieldType.SelectedItem
            $f.SelectedTypeId = if ($selectedType) { [string]$selectedType.Id } else { 'Auto' }
            $f.SelectedFormat = if ($selectedType) { [string]$selectedType.Format } else { $null }
            $f.EffectiveType = GetEffectiveTypeLabel $f
            $f.TypeStatus = GetFieldTypeStatus $f
            $f.CandidateSummary = GetFieldCandidateSummary $f

            # Update references in OutputParts and Mappings when name changed
            if ($oldName -ne $newName) {
                foreach ($p in $script:OutputParts) {
                    if ($p.Type -eq 'Field' -and $p.Value -eq $oldName) {
                        $p.Value = $newName
                        $p.Display = "$(T 'Tag_Field') $newName"
                    }
                }
                foreach ($m in $script:Mappings) {
                    $changed = $false
                    if ($m.InputField -eq $oldName) { $m.InputField = $newName; $changed = $true }
                    if ($m.OutputField -eq $oldName) { $m.OutputField = $newName; $changed = $true }
                    if ($changed) { $m.Display = "$($m.Name): $($m.InputField) → $($m.OutputField)" }
                }
                $OutputList.ItemsSource = $null
                $OutputList.ItemsSource = $script:OutputParts
                $MappingList.ItemsSource = $null
                $MappingList.ItemsSource = $script:Mappings
            }

            $FieldGrid.ItemsSource = $null
            $FieldGrid.ItemsSource = $script:Fields
            RefreshFieldSelector
            UpdateOutputExample
        }
        catch { ErrorBox (T 'Err_Field') $_ }
    })

# --- Tab 3: Transforms ---
$TransformAdd.Add_Click({
        try {
            $f = $FieldGrid.SelectedItem
            if (-not $f) { throw (T 'Err_SelFieldTrans') }
            $result = ShowTransformDialog $f
            if ($result) {
                if (-not $f.Transforms) {
                    $f.Transforms = [System.Collections.ArrayList]::new()
                }
                [void]$f.Transforms.Add($result)
                $TransformList.ItemsSource = $null
                $TransformList.ItemsSource = @($f.Transforms)
                UpdateOutputExample
            }
        }
        catch { ErrorBox (T 'Err_Transform') $_ }
    })

$TransformRemove.Add_Click({
        $f = $FieldGrid.SelectedItem
        $i = $TransformList.SelectedIndex
        if ($f -and $f.Transforms -and $i -ge 0) {
            $f.Transforms.RemoveAt($i)
            $TransformList.ItemsSource = $null
            $TransformList.ItemsSource = @($f.Transforms)
            UpdateOutputExample
        }
    })

# --- Tab 3: Mappings ---
$MappingAdd.Add_Click({
        try { AddMappingDialog; UpdateOutputExample }
        catch { ErrorBox (T 'Err_Mapping') $_ }
    })

$MappingEdit.Add_Click({
        $i = $MappingList.SelectedIndex
        if ($i -lt 0) { throw (T 'Err_SelMapping') }
        try { AddMappingDialog $script:Mappings[$i]; UpdateOutputExample }
        catch { ErrorBox (T 'Err_Mapping') $_ }
    })

$MappingRemove.Add_Click({
        $i = $MappingList.SelectedIndex
        if ($i -ge 0) {
            # Remove virtual field if no other mapping produces it
            $removedOutput = $script:Mappings[$i].OutputField
            $script:Mappings.RemoveAt($i)

            $stillUsed = $false
            foreach ($m in $script:Mappings) {
                if ($m.OutputField -eq $removedOutput) { $stillUsed = $true; break }
            }
            if (-not $stillUsed) {
                $toRemove = $null
                foreach ($f in $script:Fields) {
                    if ($f.IsVirtual -and $f.Name -eq $removedOutput) { $toRemove = $f; break }
                }
                if ($toRemove) { $script:Fields.Remove($toRemove) }
            }

            $MappingList.ItemsSource = $null
            $MappingList.ItemsSource = $script:Mappings
            $FieldGrid.ItemsSource = $null
            $FieldGrid.ItemsSource = $script:Fields
            RefreshFieldSelector
            UpdateOutputExample
        }
    })

$MappingUp.Add_Click({
        $i = $MappingList.SelectedIndex
        if ($i -gt 0) {
            $x = $script:Mappings[$i]
            $script:Mappings.RemoveAt($i)
            $script:Mappings.Insert($i - 1, $x)
            $MappingList.ItemsSource = $null
            $MappingList.ItemsSource = $script:Mappings
            $MappingList.SelectedIndex = $i - 1
        }
    })

$MappingDown.Add_Click({
        $i = $MappingList.SelectedIndex
        if ($i -ge 0 -and $i -lt $script:Mappings.Count - 1) {
            $x = $script:Mappings[$i]
            $script:Mappings.RemoveAt($i)
            $script:Mappings.Insert($i + 1, $x)
            $MappingList.ItemsSource = $null
            $MappingList.ItemsSource = $script:Mappings
            $MappingList.SelectedIndex = $i + 1
        }
    })

# --- Tab 4: Output name builder ---
$OutputAddField.Add_Click({
        try {
            $fieldName = [string]$FieldSelector.SelectedItem
            if (-not $fieldName) { throw (T 'Err_SelFieldCombo') }
            $script:OutputParts.Add([pscustomobject]@{
                    Type    = 'Field'
                    Value   = $fieldName
                    Display = "$(T 'Tag_Field') $fieldName"
                })
            $OutputList.ItemsSource = $null
            $OutputList.ItemsSource = $script:OutputParts
            UpdateOutputExample
        }
        catch { ErrorBox (T 'Title_Error') $_ }
    })

$OutputAddText.Add_Click({
        $v = $OutputText.Text
        if ($v) {
            $script:OutputParts.Add([pscustomobject]@{
                    Type    = 'Text'
                    Value   = $v
                    Display = "$(T 'Tag_Text') $v"
                })
            $OutputText.Clear()
            $OutputList.ItemsSource = $null
            $OutputList.ItemsSource = $script:OutputParts
            UpdateOutputExample
        }
    })

$OutputAddSeparator.Add_Click({
        $script:OutputParts.Add([pscustomobject]@{
                Type    = 'Text'
                Value   = '_'
                Display = "$(T 'Tag_Separator') _"
            })
        $OutputList.ItemsSource = $null
        $OutputList.ItemsSource = $script:OutputParts
        UpdateOutputExample
    })

$OutputRemove.Add_Click({
        $i = $OutputList.SelectedIndex
        if ($i -ge 0) {
            $script:OutputParts.RemoveAt($i)
            $OutputList.ItemsSource = $null
            $OutputList.ItemsSource = $script:OutputParts
            UpdateOutputExample
        }
    })

$OutputUp.Add_Click({
        $i = $OutputList.SelectedIndex
        if ($i -gt 0) {
            $x = $script:OutputParts[$i]
            $script:OutputParts.RemoveAt($i)
            $script:OutputParts.Insert($i - 1, $x)
            $OutputList.ItemsSource = $null
            $OutputList.ItemsSource = $script:OutputParts
            $OutputList.SelectedIndex = $i - 1
            UpdateOutputExample
        }
    })

$OutputDown.Add_Click({
        $i = $OutputList.SelectedIndex
        if ($i -ge 0 -and $i -lt $script:OutputParts.Count - 1) {
            $x = $script:OutputParts[$i]
            $script:OutputParts.RemoveAt($i)
            $script:OutputParts.Insert($i + 1, $x)
            $OutputList.ItemsSource = $null
            $OutputList.ItemsSource = $script:OutputParts
            $OutputList.SelectedIndex = $i + 1
            UpdateOutputExample
        }
    })

$KeepExtension.Add_Checked({ $NewExtension.IsEnabled = $false })
$KeepExtension.Add_Unchecked({ $NewExtension.IsEnabled = $true })

# --- Tab 5: Preview and execution ---
$BuildPreview.Add_Click({
        try { FullBuildPreview; SetStatus (T 'Status_PrevBuilt') }
        catch { ErrorBox (T 'Err_Preview') $_ }
    })

$PreviewFilter.Add_SelectionChanged({ RefreshPreviewGrid })

$ExportAudit.Add_Click({
        try {
            if (-not $script:PreviewRows.Count) { throw (T 'Err_BuildPrev1') }
            $dlg = New-Object Microsoft.Win32.SaveFileDialog
            $dlg.Filter = 'CSV (*.csv)|*.csv'
            $dlg.FileName = 'FileNameTransformationAudit.csv'
            if ($dlg.ShowDialog()) {
                $script:PreviewRows | Export-Csv -LiteralPath $dlg.FileName -NoTypeInformation -Encoding UTF8 -UseCulture
                SetStatus "$(T 'Status_Exported') $($dlg.FileName)"
            }
        }
        catch { ErrorBox (T 'Err_Export') $_ }
    })

$OpenLog.Add_Click({
        Start-Process notepad.exe "`"$script:LogPath`""
    })

$Execute.Add_Click({
        try { ExecuteCopy }
        catch { ErrorBox (T 'Err_Execute') $_ }
    })

# --- Tab 1: Profiles ---
$ProfileNew.Add_Click({
        $script:Fields.Clear()
        $script:Mappings.Clear()
        $script:OutputParts.Clear()
        $FieldGrid.ItemsSource = $script:Fields
        $MappingList.ItemsSource = $script:Mappings
        $OutputList.ItemsSource = $script:OutputParts
        $TransformList.ItemsSource = $null
        RefreshFieldSelector
        $script:CurrentProfileName = ''
        $CurrentProfile.Text = (T 'Txt_Unsaved')
        UpdateOutputExample
    })

$ProfileSave.Add_Click({
        try { SaveProfile }
        catch { ErrorBox (T 'Err_SaveProfile') $_ }
    })

$ProfileLoad.Add_Click({
        try {
            if (-not $ProfileList.SelectedItem) { throw (T 'Err_SelProfList') }
            LoadProfile $ProfileList.SelectedItem.Path
        }
        catch { ErrorBox (T 'Err_LoadProfile') $_ }
    })

$ProfileCopy.Add_Click({
        try {
            if (-not $ProfileList.SelectedItem) { throw (T 'Err_SelProfCopy') }
            LoadProfile $ProfileList.SelectedItem.Path
            $script:CurrentProfileName = (T 'Prefix_Copy') + $script:CurrentProfileName
            SaveProfile
        }
        catch { ErrorBox (T 'Err_CopyProfile') $_ }
    })

$ProfileDelete.Add_Click({
        try {
            if (-not $ProfileList.SelectedItem) { throw (T 'Err_SelProfDel') }
            $name = $ProfileList.SelectedItem.Name
            $confirm = [Windows.MessageBox]::Show(
                "$(T 'Msg_DelProf') '$name'?", (T 'Title_Confirm'), 'YesNo', 'Warning'
            )
            if ($confirm -eq 'Yes') {
                Remove-Item $ProfileList.SelectedItem.Path -Force
                RefreshProfiles
                SetStatus "$(T 'Status_ProfDel') $name"
            }
        }
        catch { ErrorBox (T 'Err_DelProfile') $_ }
    })

#endregion

#region Initialization
$LanguageSelector.Add_SelectionChanged({
        if ($LanguageSelector.SelectedItem) {
            $tag = $LanguageSelector.SelectedItem.Tag
            if ($tag -and $tag -ne $script:CurrentLanguage) {
                $script:Config = Set-FNTConfigLanguage -Config $script:Config -Language $tag
                $script:Config | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $script:ConfigPath -Encoding UTF8
                [Windows.MessageBox]::Show((T 'Msg_ConfirmRestart'), (T 'Title_Info'), 'OK', 'Information') | Out-Null
            }
        }
    })
$LogText.Text = "Log: $script:LogPath"
$FieldRole.SelectedIndex = 0
RefreshProfiles
UpdateOutputExample

$window.Add_Closed({ Log (T 'Log_AppClosed') })
SetStatus (T 'Status_Ready')

[void]$window.ShowDialog()
#endregion