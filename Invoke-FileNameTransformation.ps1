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

    Features:
    - Multi-language UI (Polish, English, German).
    - Saveable and reusable profiles in JSON format.
    - Pattern-based source file parsing.
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
    Requires Windows PowerShell 5.1 and an STA host. Profiles, logs, and the language configuration are stored
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
if (-not (Get-Variable -Name AppRoot -Scope Script -ErrorAction SilentlyContinue)) {
    try {
        $script:AppRoot = Split-Path -Parent $MyInvocation.MyCommand.Path -ErrorAction Stop
    }
    catch {
        $script:AppRoot = Get-Location
    }
}
$script:ConfigPath = Join-Path $script:AppRoot 'config.json'
$script:CurrentLanguage = 'PL'
$script:Config = [pscustomobject][ordered]@{
    Version         = 2
    Language        = 'PL'
    CustomTypeRules = @()
}
$coreModulePath = Join-Path $script:AppRoot 'FileNameTransformation.Core.psm1'
if (-not (Test-Path -LiteralPath $coreModulePath -PathType Leaf)) {
    throw "Missing core module: $coreModulePath"
}
Import-Module $coreModulePath -Force

if (Test-Path $script:ConfigPath) {
    try {
        Write-Verbose $script:ConfigPath
        $config = Get-Content -LiteralPath $script:ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
        $script:Config = $config
        if (-not $script:Config.PSObject.Properties['Version']) {
            $script:Config | Add-Member -NotePropertyName Version -NotePropertyValue 1
        }
        if (-not $script:Config.PSObject.Properties['CustomTypeRules']) {
            $script:Config | Add-Member -NotePropertyName CustomTypeRules -NotePropertyValue @()
        }
        if ($config.Language -in @('PL', 'EN', 'DE')) {
            $script:CurrentLanguage = $config.Language
            Write-Information "Using saved language: $script:CurrentLanguage"
        }
    }
    catch {
        Write-Error 'Failed to load config file!'
    }
}

$script:Translations = @{
    'PL' = @{
        'WinTitle'               = 'File Name Transformer'
        'Folders_Header'         = 'Foldery'
        'Source_Folder'          = 'Folder źródłowy:'
        'Btn_Browse'             = 'Wybierz...'
        'Dest_Folder'            = 'Folder wynikowy:'
        'Lbl_Language'           = 'Język:'
        'Chk_Recursive'          = 'Skanuj także podfoldery'
        'Chk_PreserveStructure'  = 'Zachowaj strukturę folderów w folderze wynikowym'
        'Tab_Profile'            = '1. Profil'
        'Saved_Profiles'         = 'Zapisane profile'
        'Btn_New'                = 'Nowy'
        'Btn_Load'               = 'Użyj'
        'Btn_Copy'               = 'Kopiuj'
        'Btn_Delete'             = 'Usuń'
        'Btn_Save_Config'        = 'Zapisz bieżącą konfigurację'
        'Title_Wizard'           = 'Kreator zmiany nazw plików'
        'Txt_ProfileHint'        = 'Pracuj kolejno na kartach 2–5. Wybierz istniejący profil, aby przywrócić konfigurację.'
        'Txt_ActiveProfile'      = 'Aktywny profil:'
        'Txt_Unsaved'            = '(niezapisana konfiguracja)'
        'Tab_Analysis'           = '2. Analiza plików'
        'Btn_Analyze'            = 'Analizuj folder'
        'Txt_Extension'          = 'Rozszerzenie:'
        'Header_DetectedStructs' = 'Wykryte struktury nazw'
        'Col_Extension'          = 'Rozszerzenie'
        'Col_Structure'          = 'Struktura'
        'Col_Files'              = 'Pliki'
        'Header_Examples'        = 'Przykłady wybranego wzorca'
        'Tab_Fields'             = '3. Pola i mapowania'
        'Header_DetectedFields'  = 'Pola wykryte w nazwie'
        'Txt_FieldHint'          = 'Zaznacz pole, aby edytować jego nazwę, rolę i transformacje. Pola oznaczone V pochodzą z mapowań.'
        'Col_Sample'             = 'Przykład'
        'Col_Type'               = 'Typ'
        'Col_Name'               = 'Nazwa'
        'Col_Role'               = 'Rola'
        'Col_Source'             = 'Źródło'
        'Header_EditField'       = 'Edycja pola'
        'Txt_FieldName'          = 'Nazwa pola:'
        'Txt_Role'               = 'Rola:'
        'Role_Value'             = 'Wartość'
        'Role_Date'              = 'Data'
        'Role_Id'                = 'Identyfikator'
        'Role_Const'             = 'Stały tekst'
        'Role_Ignore'            = 'Ignoruj'
        'Btn_ApplyField'         = 'Zastosuj zmiany pola'
        'Header_Transformations' = 'Transformacje pola'
        'Txt_TransformHint'      = 'Transformacje modyfikują wartość pola w nazwie docelowej (np. zmiana formatu daty, przycinanie tekstu, zamiana wielkości liter).'
        'Btn_AddTransform'       = 'Dodaj transformację'
        'Btn_Remove'             = 'Usuń'
        'Header_Mappings'        = 'Mapowania (zamiana wartości z pliku CSV/TXT)'
        'Txt_MappingHint'        = 'Mapowanie zamienia wartość pola na inną wartość z pliku zewnętrznego. Tworzy nowe pole wirtualne dostępne w nazwie docelowej.'
        'Btn_AddMapping'         = 'Dodaj mapowanie'
        'Btn_Edit'               = 'Edytuj'
        'Title_EditMapping'      = 'Edytuj mapowanie'
        'Btn_Up'                 = 'W górę'
        'Btn_Down'               = 'W dół'
        'Tab_DestName'           = '4. Nazwa docelowa'
        'Header_DestElements'    = 'Elementy nazwy docelowej'
        'Btn_AddField'           = 'Dodaj pole'
        'Btn_AddText'            = 'Dodaj tekst'
        'Btn_AddSep'             = 'Dodaj _'
        'Btn_RemoveElement'      = 'Usuń element'
        'Header_Settings'        = 'Ustawienia i podgląd na żywo'
        'Chk_KeepExt'            = 'Zachowaj oryginalne rozszerzenie'
        'Txt_NewExt'             = 'Nowe rozszerzenie:'
        'Txt_PreviewTitle'       = 'Podgląd nazwy (pierwszy plik):'
        'Tab_Preview'            = '5. Podgląd i wykonanie'
        'Btn_BuildPreview'       = 'Utwórz podgląd'
        'Cbo_All'                = 'Wszystkie'
        'Cbo_Errors'             = 'Tylko błędy'
        'Cbo_Ready'              = 'Gotowe'
        'Btn_Export'             = 'Eksportuj raport CSV'
        'Btn_OpenLog'            = 'Otwórz log'
        'Btn_Execute'            = 'Wykonaj'
        'Lbl_ExecutionMode'      = 'Akcja:'
        'Action_Copy'            = 'Kopiuj'
        'Action_Move'            = 'Przenieś'
        'Col_SourceFile'         = 'Plik źródłowy'
        'Col_DestName'           = 'Nazwa docelowa'
        'Col_Status'             = 'Status'
        'Col_Details'            = 'Szczegóły'
        'Msg_ConfirmRestart'     = 'Język został zmieniony. Uruchom aplikację ponownie, aby w pełni zastosować nowy język.'
        'Title_Info'             = 'Informacja'
        'Msg_SelectSource'       = 'Wybierz folder źródłowy'
        'Msg_SelectDest'         = 'Wybierz folder wynikowy'
        'Log_AppClosed'          = 'Aplikacja zamknięta.'
        'Status_Ready'           = 'Gotowy. Wybierz folder źródłowy i rozpocznij analizę na karcie 2.'
        'Err_Analysis'           = 'Błąd analizy'
        'Err_Field'              = 'Błąd pola'
        'Err_Transform'          = 'Błąd transformacji'
        'Err_Mapping'            = 'Błąd mapowania'
        'Err_Preview'            = 'Błąd podglądu'
        'Err_Export'             = 'Błąd eksportu'
        'Err_Execute'            = 'Błąd wykonania'
        'Err_SaveProfile'        = 'Błąd zapisu profilu'
        'Err_LoadProfile'        = 'Błąd wczytania profilu'
        'Err_CopyProfile'        = 'Błąd kopiowania profilu'
        'Err_DelProfile'         = 'Błąd usuwania profilu'
        'Title_Error'            = 'Błąd'
        'Title_Warning'          = 'Ostrzeżenie'
        'Title_Confirm'          = 'Potwierdzenie'
        'Title_SaveProfile'      = 'Zapis profilu'
        'Type_Date_1'            = 'Data (yyyyMMdd)'
        'Type_Date_2'            = 'Data (yyyy-MM-dd)'
        'Type_Date_3'            = 'Data (dd-MM-yyyy)'
        'Type_Date_4'            = 'Data (yyyy-MM)'
        'Type_Date_5'            = 'Data (MM-yyyy)'
        'Type_Date_6'            = 'Data (YYMMDD lub DDMMYY)'
        'Type_Date_7'            = 'Data (MM-DD lub DD-MM)'
        'Type_Date_8'            = 'Data (YYYY-MM-DDTHH:MM:SS lub YYYY-MM-DDTHH:MM:SSZ)'
        'Type_Num'               = 'Liczba'
        'Type_Text'              = 'Tekst'
        'Type_Integer'           = 'Liczba całkowita'
        'Type_Decimal'           = 'Liczba dziesiętna'
        'Type_DateTime'          = 'Data/czas'
        'Type_Guid'              = 'GUID'
        'Type_Version'           = 'Wersja'
        'Type_Ambiguous'         = 'Niejednoznaczny'
        'Type_Auto'              = 'Automatyczny'
        'Txt_DataType'           = 'Typ danych:'
        'Err_ResolveAmbiguous'    = 'Wybierz jednoznaczny typ danych dla pola:'
        'Hint_SelectField'       = 'Wybierz pole, nadaj mu nazwę biznesową i dodaj transformacje lub mapowania.'
        'Hint_AddElements'       = 'Dodaj elementy nazwy docelowej za pomocą przycisków poniżej.'
        'Prefix_Source'          = 'Źródło'
        'Prefix_Result'          = 'Wynik'
        'Preview_Unavailable'    = 'Podgląd niedostępny'
        'Err_SrcNotExist'        = 'Wskaż istniejący folder źródłowy.'
        'Err_NoFiles'            = 'Nie znaleziono plików.'
        'Msg_Files'              = 'Pliki'
        'Msg_Structs'            = 'wykryte struktury'
        'Msg_Errors'             = 'błędy'
        'Name_Date'              = 'Data'
        'Name_Text'              = 'Tekst'
        'Name_Field'             = 'Pole'
        'Src_Name'               = 'Nazwa'
        'Src_Mapping'            = 'Mapowanie'
        'Val_Mapping'            = '(mapowanie)'
        'Title_AddTransform'     = 'Dodaj transformację'
        'Txt_TransformType'      = 'Typ transformacji:'
        'Tr_Substring'           = 'Wycinanie tekstu (Substring)'
        'Tr_DateFormat'          = 'Format daty'
        'Tr_Replace'             = 'Zamiana tekstu'
        'Tr_Case'                = 'Zmiana wielkości liter'
        'Tr_Pad'                 = 'Dopełnienie znakami'
        'Tr_PosStart'            = 'Pozycja startowa (0 = początek):'
        'Tr_CharCount'           = 'Liczba znaków:'
        'Tr_FmtIn'               = 'Format wejściowy:'
        'Tr_FmtOut'              = 'Format wyjściowy:'
        'Tr_Search'              = 'Szukany tekst:'
        'Tr_NewTxt'              = 'Nowy tekst (pusty = usuń):'
        'Tr_Mode'                = 'Tryb:'
        'Tr_Upper'               = 'WIELKIE LITERY'
        'Tr_Lower'               = 'małe litery'
        'Tr_Title'               = 'Pierwsza Wielka'
        'Tr_Side'                = 'Strona:'
        'Tr_Left'                = 'Z lewej'
        'Tr_Right'               = 'Z prawej'
        'Tr_PadChar'             = 'Znak dopełnienia:'
        'Tr_TargetLen'           = 'Docelowa długość:'
        'Btn_Add'                = 'Dodaj'
        'Err_CharCount'          = 'Liczba znaków musi być > 0.'
        'Err_BothFmt'            = 'Podaj oba formaty.'
        'Err_SearchTxt'          = 'Podaj szukany tekst.'
        'Err_SelMode'            = 'Wybierz tryb.'
        'Err_SelSide'            = 'Wybierz stronę.'
        'Err_PadChar'            = 'Podaj znak dopełnienia.'
        'Err_TargetLen'          = 'Długość musi być > 0.'
        'Err_SelTransform'       = 'Wybierz typ transformacji.'
        'Disp_Substr'            = 'Wycinanie: od pozycji'
        'Disp_Chars'             = 'znaków'
        'Disp_Date'              = 'Data:'
        'Disp_Replace'           = 'Zamiana:'
        'Disp_Case'              = 'Wielkość liter:'
        'Disp_Pad'               = 'Dopełnienie:'
        'Disp_To'                = 'do'
        'Title_AddMapping'       = 'Dodaj mapowanie'
        'Lbl_MapName'            = 'Nazwa mapowania'
        'Lbl_MapIn'              = 'Pole wejściowe'
        'Lbl_MapOut'             = 'Nazwa pola wynikowego'
        'Lbl_MapFile'            = 'Plik danych (CSV/TXT)'
        'Lbl_MapKey'             = 'Kolumna klucza'
        'Lbl_MapVal'             = 'Kolumna wartości'
        'Err_ReadHeaders'        = 'Nie udało się odczytać nagłówków:'
        'Title_FileErr'          = 'Błąd pliku'
        'Err_FillFields'         = 'Uzupełnij wszystkie pola.'
        'Title_MissingData'      = 'Brak danych'
        'Err_SelMapping'         = 'Wybierz mapowanie do edycji.'
        'Err_SelectPattern'      = 'Najpierw wybierz wzorzec na karcie 2 (Analiza plików).'
        'Err_AddDest'            = 'Dodaj elementy nazwy docelowej na karcie 4.'
        'Err_DestNotExist'       = 'Wskaż istniejący folder wynikowy.'
        'Err_MissMapFile'        = 'Brak pliku mapowania:'
        'Err_DupKey'             = 'Duplikat klucza'
        'Err_InMap'              = 'w mapowaniu'
        'Err_MissMap'            = 'Brak mapowania'
        'Err_MissField'          = 'Brak pola'
        'Err_InvalidChars'       = 'Nazwa docelowa zawiera niedozwolone znaki.'
        'Err_DupDest'            = 'Zduplikowana ścieżka docelowa'
        'Err_FileExists'         = 'Plik docelowy już istnieje'
        'Err_Blocked'            = 'Operacja zablokowana:'
        'Err_FixPrev'            = 'błędów w podglądzie. Napraw je przed kopiowaniem.'
        'Msg_CopyCount'          = 'Skopiować'
        'Msg_CopyOrig'           = 'plików? Oryginały nie zostaną zmienione.'
        'Msg_MoveOrig'           = 'plików? Oryginały zostaną przeniesione.'
        'Log_Copied'             = 'Skopiowano:'
        'Log_Moved'              = 'Przeniesiono:'
        'Log_CopyErr'            = 'Błąd kopiowania:'
        'Status_Copied'          = 'Zakończono. Skopiowano:'
        'Status_Moved'           = 'Zakończono. Przeniesiono:'
        'Lbl_ProfileName'        = 'Nazwa profilu:'
        'Status_ProfSaved'       = 'Profil zapisany:'
        'Status_ProfLoaded'      = 'Profil wczytany:'
        'Status_AnalysisDone'    = 'Analiza zakończona.'
        'Err_SelFieldTab'        = 'Zaznacz pole w tabeli.'
        'Err_FieldName'          = 'Podaj nazwę pola.'
        'Err_SelFieldTrans'      = 'Zaznacz pole, do którego chcesz dodać transformację.'
        'Err_SelFieldCombo'      = 'Wybierz pole z listy rozwijanej.'
        'Tag_Field'              = '[Pole]'
        'Tag_Text'               = '[Tekst]'
        'Tag_Separator'          = '[Separator]'
        'Status_PrevBuilt'       = 'Podgląd utworzony.'
        'Err_BuildPrev1'         = 'Najpierw utwórz podgląd.'
        'Status_Exported'        = 'Raport wyeksportowany:'
        'Err_SelProfList'        = 'Wybierz profil z listy.'
        'Err_SelProfCopy'        = 'Wybierz profil do skopiowania.'
        'Prefix_Copy'            = 'Kopia '
        'Err_SelProfDel'         = 'Wybierz profil do usunięcia.'
        'Msg_DelProf'            = 'Usunąć profil'
        'Status_ProfDel'         = 'Profil usunięty:'
        'Chk_EnforcePattern'     = 'Wymuś wzorzec na pozostałych plikach'
        'Txt_TokenRegex'         = 'Wzorzec Regex:'
        'Tip_TokenRegex'         = 'Domyślnie: (?&lt;value&gt;[^_\-\s]+)|(?&lt;sep&gt;[_\-\s]+)'
        'Tip_TokenRegexLabel'    = 'Regex musi dopasować każdy segment nazwy. Użyj (?&lt;sep&gt;...) dla separatorów; pozostałe dopasowania są wartościami. Przykład: (?&lt;value&gt;[^_\-\s]+)|(?&lt;sep&gt;[_\-\s]+).'
    }
    'EN' = @{
        'WinTitle'               = 'File Name Transformer'
        'Folders_Header'         = 'Folders'
        'Source_Folder'          = 'Source folder:'
        'Btn_Browse'             = 'Browse...'
        'Dest_Folder'            = 'Destination folder:'
        'Lbl_Language'           = 'Language:'
        'Chk_Recursive'          = 'Scan subfolders'
        'Chk_PreserveStructure'  = 'Preserve folder structure in destination folder'
        'Tab_Profile'            = '1. Profile'
        'Saved_Profiles'         = 'Saved profiles'
        'Btn_New'                = 'New'
        'Btn_Load'               = 'Load'
        'Btn_Copy'               = 'Copy'
        'Btn_Delete'             = 'Delete'
        'Btn_Save_Config'        = 'Save current configuration'
        'Title_Wizard'           = 'File Name Transformer Wizard'
        'Txt_ProfileHint'        = 'Work sequentially on tabs 2-5. Select an existing profile to restore configuration.'
        'Txt_ActiveProfile'      = 'Active profile:'
        'Txt_Unsaved'            = '(unsaved configuration)'
        'Tab_Analysis'           = '2. File analysis'
        'Btn_Analyze'            = 'Analyze folder'
        'Txt_Extension'          = 'Extension:'
        'Header_DetectedStructs' = 'Detected name structures'
        'Col_Extension'          = 'Extension'
        'Col_Structure'          = 'Structure'
        'Col_Files'              = 'Files'
        'Header_Examples'        = 'Examples of selected pattern'
        'Tab_Fields'             = '3. Fields and mappings'
        'Header_DetectedFields'  = 'Fields detected in name'
        'Txt_FieldHint'          = 'Select a field to edit its name, role and transformations. Fields marked with V come from mappings.'
        'Col_Sample'             = 'Sample'
        'Col_Type'               = 'Type'
        'Col_Name'               = 'Name'
        'Col_Role'               = 'Role'
        'Col_Source'             = 'Source'
        'Header_EditField'       = 'Edit field'
        'Txt_FieldName'          = 'Field name:'
        'Txt_Role'               = 'Role:'
        'Role_Value'             = 'Value'
        'Role_Date'              = 'Date'
        'Role_Id'                = 'Identifier'
        'Role_Const'             = 'Constant text'
        'Role_Ignore'            = 'Ignore'
        'Btn_ApplyField'         = 'Apply field changes'
        'Header_Transformations' = 'Field transformations'
        'Txt_TransformHint'      = 'Transformations modify field value in destination name (e.g., date format change, text trimming, case change).'
        'Btn_AddTransform'       = 'Add transformation'
        'Btn_Remove'             = 'Remove'
        'Header_Mappings'        = 'Mappings (value replacement from CSV/TXT)'
        'Txt_MappingHint'        = 'Mapping replaces a field value with another value from an external file. It creates a new virtual field available in destination name.'
        'Btn_AddMapping'         = 'Add mapping'
        'Btn_Edit'               = 'Edit'
        'Title_EditMapping'      = 'Edit mapping'
        'Btn_Up'                 = 'Move up'
        'Btn_Down'               = 'Move down'
        'Tab_DestName'           = '4. Destination name'
        'Header_DestElements'    = 'Destination name elements'
        'Btn_AddField'           = 'Add field'
        'Btn_AddText'            = 'Add text'
        'Btn_AddSep'             = 'Add _'
        'Btn_RemoveElement'      = 'Remove element'
        'Header_Settings'        = 'Settings and live preview'
        'Chk_KeepExt'            = 'Keep original extension'
        'Txt_NewExt'             = 'New extension:'
        'Txt_PreviewTitle'       = 'Name preview (first file):'
        'Tab_Preview'            = '5. Preview and execution'
        'Btn_BuildPreview'       = 'Build preview'
        'Cbo_All'                = 'All'
        'Cbo_Errors'             = 'Errors only'
        'Cbo_Ready'              = 'Ready'
        'Btn_Export'             = 'Export CSV report'
        'Btn_OpenLog'            = 'Open log'
        'Btn_Execute'            = 'Execute'
        'Lbl_ExecutionMode'      = 'Action:'
        'Action_Copy'            = 'Copy'
        'Action_Move'            = 'Move'
        'Col_SourceFile'         = 'Source file'
        'Col_DestName'           = 'Destination name'
        'Col_Status'             = 'Status'
        'Col_Details'            = 'Details'
        'Msg_ConfirmRestart'     = 'Language changed. Please restart the application to fully apply the new language.'
        'Title_Info'             = 'Information'
        'Msg_SelectSource'       = 'Select source folder'
        'Msg_SelectDest'         = 'Select destination folder'
        'Log_AppClosed'          = 'Application closed.'
        'Status_Ready'           = 'Ready. Select source folder and start analysis on tab 2.'
        'Err_Analysis'           = 'Analysis error'
        'Err_Field'              = 'Field error'
        'Err_Transform'          = 'Transformation error'
        'Err_Mapping'            = 'Mapping error'
        'Err_Preview'            = 'Preview error'
        'Err_Export'             = 'Export error'
        'Err_Execute'            = 'Execution error'
        'Err_SaveProfile'        = 'Profile save error'
        'Err_LoadProfile'        = 'Profile load error'
        'Err_CopyProfile'        = 'Profile copy error'
        'Err_DelProfile'         = 'Profile delete error'
        'Title_Error'            = 'Error'
        'Title_Warning'          = 'Warning'
        'Title_Confirm'          = 'Confirmation'
        'Title_SaveProfile'      = 'Save profile'
        'Type_Date_1'            = 'Date (yyyyMMdd)'
        'Type_Date_2'            = 'Date (yyyy-MM-dd)'
        'Type_Date_3'            = 'Date (dd-MM-yyyy)'
        'Type_Date_4'            = 'Date (yyyy-MM)'
        'Type_Date_5'            = 'Date (MM-yyyy)'
        'Type_Date_6'            = 'Date (YYMMDD or DDMMYY)'
        'Type_Date_7'            = 'Date (MM-DD or DD-MM)'
        'Type_Date_8'            = 'Date (YYYY-MM-DDTHH:MM:SS or YYYY-MM-DDTHH:MM:SSZ)'
        'Type_Num'               = 'Number'
        'Type_Text'              = 'Text'
        'Type_Integer'           = 'Integer'
        'Type_Decimal'           = 'Decimal'
        'Type_DateTime'          = 'Date/time'
        'Type_Guid'              = 'GUID'
        'Type_Version'           = 'Version'
        'Type_Ambiguous'         = 'Ambiguous'
        'Type_Auto'              = 'Auto'
        'Txt_DataType'           = 'Data type:'
        'Err_ResolveAmbiguous'    = 'Select an unambiguous data type for field:'
        'Hint_SelectField'       = 'Select a field, give it a business name, and add transformations or mappings.'
        'Hint_AddElements'       = 'Add destination name elements using buttons below.'
        'Prefix_Source'          = 'Source'
        'Prefix_Result'          = 'Result'
        'Preview_Unavailable'    = 'Preview unavailable'
        'Err_SrcNotExist'        = 'Specify an existing source folder.'
        'Err_NoFiles'            = 'No files found.'
        'Msg_Files'              = 'Files'
        'Msg_Structs'            = 'detected structures'
        'Msg_Errors'             = 'errors'
        'Name_Date'              = 'Date'
        'Name_Text'              = 'Text'
        'Name_Field'             = 'Field'
        'Src_Name'               = 'Name'
        'Src_Mapping'            = 'Mapping'
        'Val_Mapping'            = '(mapping)'
        'Title_AddTransform'     = 'Add transformation'
        'Txt_TransformType'      = 'Transformation type:'
        'Tr_Substring'           = 'Text trimming (Substring)'
        'Tr_DateFormat'          = 'Date format'
        'Tr_Replace'             = 'Text replacement'
        'Tr_Case'                = 'Case change'
        'Tr_Pad'                 = 'Padding'
        'Tr_PosStart'            = 'Start position (0 = start):'
        'Tr_CharCount'           = 'Number of characters:'
        'Tr_FmtIn'               = 'Input format:'
        'Tr_FmtOut'              = 'Output format:'
        'Tr_Search'              = 'Search text:'
        'Tr_NewTxt'              = 'New text (empty = remove):'
        'Tr_Mode'                = 'Mode:'
        'Tr_Upper'               = 'UPPERCASE'
        'Tr_Lower'               = 'lowercase'
        'Tr_Title'               = 'Title Case'
        'Tr_Side'                = 'Side:'
        'Tr_Left'                = 'Left'
        'Tr_Right'               = 'Right'
        'Tr_PadChar'             = 'Pad character:'
        'Tr_TargetLen'           = 'Target length:'
        'Btn_Add'                = 'Add'
        'Err_CharCount'          = 'Number of characters must be > 0.'
        'Err_BothFmt'            = 'Provide both formats.'
        'Err_SearchTxt'          = 'Provide search text.'
        'Err_SelMode'            = 'Select mode.'
        'Err_SelSide'            = 'Select side.'
        'Err_PadChar'            = 'Provide pad character.'
        'Err_TargetLen'          = 'Length must be > 0.'
        'Err_SelTransform'       = 'Select transformation type.'
        'Disp_Substr'            = 'Trimming: from pos'
        'Disp_Chars'             = 'chars'
        'Disp_Date'              = 'Date:'
        'Disp_Replace'           = 'Replace:'
        'Disp_Case'              = 'Case:'
        'Disp_Pad'               = 'Pad:'
        'Disp_To'                = 'to'
        'Title_AddMapping'       = 'Add mapping'
        'Lbl_MapName'            = 'Mapping name'
        'Lbl_MapIn'              = 'Input field'
        'Lbl_MapOut'             = 'Output field name'
        'Lbl_MapFile'            = 'Data file (CSV/TXT)'
        'Lbl_MapKey'             = 'Key column'
        'Lbl_MapVal'             = 'Value column'
        'Err_ReadHeaders'        = 'Failed to read headers:'
        'Title_FileErr'          = 'File error'
        'Err_FillFields'         = 'Fill in all fields.'
        'Title_MissingData'      = 'Missing data'
        'Err_SelectPattern'      = 'First select a pattern on tab 2 (File analysis).'
        'Err_AddDest'            = 'Add destination name elements on tab 4.'
        'Err_DestNotExist'       = 'Specify an existing destination folder.'
        'Err_MissMapFile'        = 'Missing mapping file:'
        'Err_DupKey'             = 'Duplicate key'
        'Err_InMap'              = 'in mapping'
        'Err_MissMap'            = 'Missing mapping'
        'Err_MissField'          = 'Missing field'
        'Err_InvalidChars'       = 'Destination name contains invalid characters.'
        'Err_DupDest'            = 'Duplicate destination path'
        'Err_FileExists'         = 'Destination file already exists'
        'Err_Blocked'            = 'Operation blocked:'
        'Err_FixPrev'            = 'errors in preview. Fix them before copying.'
        'Msg_CopyCount'          = 'Copy'
        'Msg_CopyOrig'           = 'files? Originals will not be modified.'
        'Log_Copied'             = 'Copied:'
        'Log_CopyErr'            = 'Copy error:'
        'Status_Copied'          = 'Finished. Copied:'
        'Lbl_ProfileName'        = 'Profile name:'
        'Status_ProfSaved'       = 'Profile saved:'
        'Status_ProfLoaded'      = 'Profile loaded:'
        'Status_AnalysisDone'    = 'Analysis complete.'
        'Err_SelFieldTab'        = 'Select a field in the table.'
        'Err_FieldName'          = 'Provide field name.'
        'Err_SelFieldTrans'      = 'Select a field to add transformation to.'
        'Err_SelFieldCombo'      = 'Select a field from dropdown.'
        'Tag_Field'              = '[Field]'
        'Tag_Text'               = '[Text]'
        'Tag_Separator'          = '[Separator]'
        'Status_PrevBuilt'       = 'Preview built.'
        'Err_BuildPrev1'         = 'Build preview first.'
        'Status_Exported'        = 'Report exported:'
        'Err_SelProfList'        = 'Select a profile from list.'
        'Err_SelProfCopy'        = 'Select a profile to copy.'
        'Prefix_Copy'            = 'Copy of '
        'Err_SelProfDel'         = 'Select a profile to delete.'
        'Msg_DelProf'            = 'Delete profile'
        'Status_ProfDel'         = 'Profile deleted:'
        'Chk_EnforcePattern'     = 'Enforce pattern on other files'
        'Txt_TokenRegex'         = 'Regex Pattern:'
        'Tip_TokenRegex'         = 'Default: (?&lt;value&gt;[^_\-\s]+)|(?&lt;sep&gt;[_\-\s]+)'
        'Tip_TokenRegexLabel'    = 'The regex must match every filename segment. Use (?&lt;sep&gt;...) for separators; all other matches are values. Example: (?&lt;value&gt;[^_\-\s]+)|(?&lt;sep&gt;[_\-\s]+).'
    }
    'DE' = @{
        'WinTitle'               = 'File Name Transformer'
        'Folders_Header'         = 'Ordner'
        'Source_Folder'          = 'Quellordner:'
        'Btn_Browse'             = 'Durchsuchen...'
        'Dest_Folder'            = 'Zielordner:'
        'Lbl_Language'           = 'Sprache:'
        'Chk_Recursive'          = 'Unterordner scannen'
        'Chk_PreserveStructure'  = 'Ordnerstruktur im Zielordner beibehalten'
        'Tab_Profile'            = '1. Profil'
        'Saved_Profiles'         = 'Gespeicherte Profile'
        'Btn_New'                = 'Neu'
        'Btn_Load'               = 'Laden'
        'Btn_Copy'               = 'Kopieren'
        'Btn_Delete'             = 'Löschen'
        'Btn_Save_Config'        = 'Aktuelle Konfiguration speichern'
        'Title_Wizard'           = 'Dateinamen-Transformator-Assistent'
        'Txt_ProfileHint'        = 'Arbeiten Sie nacheinander an den Registerkarten 2-5. Wählen Sie ein vorhandenes Profil aus, um die Konfiguration wiederherzustellen.'
        'Txt_ActiveProfile'      = 'Aktives Profil:'
        'Txt_Unsaved'            = '(ungespeicherte Konfiguration)'
        'Tab_Analysis'           = '2. Dateianalyse'
        'Btn_Analyze'            = 'Ordner analysieren'
        'Txt_Extension'          = 'Erweiterung:'
        'Header_DetectedStructs' = 'Erkannte Namensstrukturen'
        'Col_Extension'          = 'Erw.'
        'Col_Structure'          = 'Struktur'
        'Col_Files'              = 'Dateien'
        'Header_Examples'        = 'Beispiele für das ausgewählte Muster'
        'Tab_Fields'             = '3. Felder und Zuordnungen'
        'Header_DetectedFields'  = 'Im Namen erkannte Felder'
        'Txt_FieldHint'          = 'Wählen Sie ein Feld aus, um seinen Namen, seine Rolle und seine Transformationen zu bearbeiten. Mit V markierte Felder stammen aus Zuordnungen.'
        'Col_Sample'             = 'Beispiel'
        'Col_Type'               = 'Typ'
        'Col_Name'               = 'Name'
        'Col_Role'               = 'Rolle'
        'Col_Source'             = 'Quelle'
        'Header_EditField'       = 'Feld bearbeiten'
        'Txt_FieldName'          = 'Feldname:'
        'Txt_Role'               = 'Rolle:'
        'Role_Value'             = 'Wert'
        'Role_Date'              = 'Datum'
        'Role_Id'                = 'Identifikator'
        'Role_Const'             = 'Konstanter Text'
        'Role_Ignore'            = 'Ignorieren'
        'Btn_ApplyField'         = 'Feldänderungen anwenden'
        'Header_Transformations' = 'Feldtransformationen'
        'Txt_TransformHint'      = 'Transformationen ändern den Feldwert im Zielnamen (z. B. Änderung des Datumsformats, Textkürzung, Groß-/Kleinschreibung).'
        'Btn_AddTransform'       = 'Transformation hinzufügen'
        'Btn_Remove'             = 'Entfernen'
        'Header_Mappings'        = 'Zuordnungen (Wertersatz aus CSV/TXT)'
        'Txt_MappingHint'        = 'Die Zuordnung ersetzt einen Feldwert durch einen anderen Wert aus einer externen Datei. Es wird ein neues virtuelles Feld erstellt.'
        'Btn_AddMapping'         = 'Zuordnung hinzufügen'
        'Btn_Edit'               = 'Bearbeiten'
        'Title_EditMapping'      = 'Zuordnung bearbeiten'
        'Btn_Up'                 = 'Nach oben'
        'Btn_Down'               = 'Nach unten'
        'Tab_DestName'           = '4. Zielname'
        'Header_DestElements'    = 'Elemente des Zielnamens'
        'Btn_AddField'           = 'Feld hinzufügen'
        'Btn_AddText'            = 'Text hinzufügen'
        'Btn_AddSep'             = '_ hinzufügen'
        'Btn_RemoveElement'      = 'Element entfernen'
        'Header_Settings'        = 'Einstellungen und Live-Vorschau'
        'Chk_KeepExt'            = 'Original-Erweiterung beibehalten'
        'Txt_NewExt'             = 'Neue Erweiterung:'
        'Txt_PreviewTitle'       = 'Namensvorschau (erste Datei):'
        'Tab_Preview'            = '5. Vorschau und Ausführung'
        'Btn_BuildPreview'       = 'Vorschau erstellen'
        'Cbo_All'                = 'Alle'
        'Cbo_Errors'             = 'Nur Fehler'
        'Cbo_Ready'              = 'Bereit'
        'Btn_Export'             = 'CSV-Bericht exportieren'
        'Btn_OpenLog'            = 'Protokoll öffnen'
        'Btn_Execute'            = 'Ausführen'
        'Lbl_ExecutionMode'      = 'Aktion:'
        'Action_Copy'            = 'Kopieren'
        'Action_Move'            = 'Verschieben'
        'Col_SourceFile'         = 'Quelldatei'
        'Col_DestName'           = 'Zielname'
        'Col_Status'             = 'Status'
        'Col_Details'            = 'Details'
        'Msg_ConfirmRestart'     = 'Sprache geändert. Bitte starten Sie die Anwendung neu, um die neue Sprache vollständig anzuwenden.'
        'Title_Info'             = 'Information'
        'Msg_SelectSource'       = 'Quellordner auswählen'
        'Msg_SelectDest'         = 'Zielordner auswählen'
        'Log_AppClosed'          = 'Anwendung geschlossen.'
        'Status_Ready'           = 'Bereit. Quellordner auswählen und Analyse auf Registerkarte 2 starten.'
        'Err_Analysis'           = 'Analysefehler'
        'Err_Field'              = 'Feldfehler'
        'Err_Transform'          = 'Transformationsfehler'
        'Err_Mapping'            = 'Zuordnungsfehler'
        'Err_Preview'            = 'Vorschaufehler'
        'Err_Export'             = 'Exportfehler'
        'Err_Execute'            = 'Ausführungsfehler'
        'Err_SaveProfile'        = 'Profil-Speicherfehler'
        'Err_LoadProfile'        = 'Profil-Ladefehler'
        'Err_CopyProfile'        = 'Profil-Kopierfehler'
        'Err_DelProfile'         = 'Profil-Löschfehler'
        'Title_Error'            = 'Fehler'
        'Title_Warning'          = 'Warnung'
        'Title_Confirm'          = 'Bestätigung'
        'Title_SaveProfile'      = 'Profil speichern'
        'Type_Date_1'            = 'Datum (yyyyMMdd)'
        'Type_Date_2'            = 'Datum (yyyy-MM-dd)'
        'Type_Date_3'            = 'Datum (dd-MM-yyyy)'
        'Type_Date_4'            = 'Datum (yyyy-MM)'
        'Type_Date_5'            = 'Datum (MM-yyyy)'
        'Type_Date_6'            = 'Datum (YYMMDD oder DDMMYY)'
        'Type_Date_7'            = 'Datum (MM-DD oder DD-MM)'
        'Type_Date_8'            = 'Datum (YYYY-MM-DDTHH:MM:SS oder YYYY-MM-DDTHH:MM:SSZ)'
        'Type_Num'               = 'Zahl'
        'Type_Text'              = 'Text'
        'Type_Integer'           = 'Ganzzahl'
        'Type_Decimal'           = 'Dezimalzahl'
        'Type_DateTime'          = 'Datum/Uhrzeit'
        'Type_Guid'              = 'GUID'
        'Type_Version'           = 'Version'
        'Type_Ambiguous'         = 'Mehrdeutig'
        'Type_Auto'              = 'Automatisch'
        'Txt_DataType'           = 'Datentyp:'
        'Err_ResolveAmbiguous'    = 'Wählen Sie einen eindeutigen Datentyp für das Feld:'
        'Hint_SelectField'       = 'Wählen Sie ein Feld aus, geben Sie ihm einen Geschäftsnamen und fügen Sie Transformationen oder Zuordnungen hinzu.'
        'Hint_AddElements'       = 'Fügen Sie Zielnamenselemente mit den Schaltflächen unten hinzu.'
        'Prefix_Source'          = 'Quelle'
        'Prefix_Result'          = 'Ergebnis'
        'Preview_Unavailable'    = 'Vorschau nicht verfügbar'
        'Err_SrcNotExist'        = 'Geben Sie einen vorhandenen Quellordner an.'
        'Err_NoFiles'            = 'Keine Dateien gefunden.'
        'Msg_Files'              = 'Dateien'
        'Msg_Structs'            = 'erkannte Strukturen'
        'Msg_Errors'             = 'Fehler'
        'Name_Date'              = 'Datum'
        'Name_Text'              = 'Text'
        'Name_Field'             = 'Feld'
        'Src_Name'               = 'Name'
        'Src_Mapping'            = 'Zuordnung'
        'Val_Mapping'            = '(Zuordnung)'
        'Title_AddTransform'     = 'Transformation hinzufügen'
        'Txt_TransformType'      = 'Transformationstyp:'
        'Tr_Substring'           = 'Textkürzung (Substring)'
        'Tr_DateFormat'          = 'Datumsformat'
        'Tr_Replace'             = 'Textersatz'
        'Tr_Case'                = 'Groß-/Kleinschreibung'
        'Tr_Pad'                 = 'Auffüllen'
        'Tr_PosStart'            = 'Startposition (0 = Anfang):'
        'Tr_CharCount'           = 'Anzahl der Zeichen:'
        'Tr_FmtIn'               = 'Eingabeformat:'
        'Tr_FmtOut'              = 'Ausgabeformat:'
        'Tr_Search'              = 'Suchtext:'
        'Tr_NewTxt'              = 'Neuer Text (leer = entfernen):'
        'Tr_Mode'                = 'Modus:'
        'Tr_Upper'               = 'GROSSBUCHSTABEN'
        'Tr_Lower'               = 'kleinbuchstaben'
        'Tr_Title'               = 'Erster Buchstabe groß'
        'Tr_Side'                = 'Seite:'
        'Tr_Left'                = 'Links'
        'Tr_Right'               = 'Rechts'
        'Tr_PadChar'             = 'Füllzeichen:'
        'Tr_TargetLen'           = 'Ziellänge:'
        'Btn_Add'                = 'Hinzufügen'
        'Err_CharCount'          = 'Die Anzahl der Zeichen muss > 0 sein.'
        'Err_BothFmt'            = 'Geben Sie beide Formate an.'
        'Err_SearchTxt'          = 'Geben Sie den Suchtext an.'
        'Err_SelMode'            = 'Wählen Sie den Modus.'
        'Err_SelSide'            = 'Wählen Sie die Seite.'
        'Err_PadChar'            = 'Geben Sie das Füllzeichen an.'
        'Err_TargetLen'          = 'Länge muss > 0 sein.'
        'Err_SelTransform'       = 'Wählen Sie den Transformationstyp.'
        'Disp_Substr'            = 'Kürzen: ab Pos'
        'Disp_Chars'             = 'Zeichen'
        'Disp_Date'              = 'Datum:'
        'Disp_Replace'           = 'Ersatz:'
        'Disp_Case'              = 'Gr./Kl.:'
        'Disp_Pad'               = 'Auffüllen:'
        'Disp_To'                = 'bis'
        'Title_AddMapping'       = 'Zuordnung hinzufügen'
        'Lbl_MapName'            = 'Zuordnungsname'
        'Lbl_MapIn'              = 'Eingabefeld'
        'Lbl_MapOut'             = 'Ausgabefeldname'
        'Lbl_MapFile'            = 'Datendatei (CSV/TXT)'
        'Lbl_MapKey'             = 'Schlüsselspalte'
        'Lbl_MapVal'             = 'Wertspalte'
        'Err_ReadHeaders'        = 'Fehler beim Lesen der Kopfzeilen:'
        'Title_FileErr'          = 'Dateifehler'
        'Err_FillFields'         = 'Füllen Sie alle Felder aus.'
        'Title_MissingData'      = 'Fehlende Daten'
        'Err_SelectPattern'      = 'Wählen Sie zunächst ein Muster auf Registerkarte 2 (Dateianalyse) aus.'
        'Err_AddDest'            = 'Fügen Sie Zielnamenselemente auf Registerkarte 4 hinzu.'
        'Err_DestNotExist'       = 'Geben Sie einen vorhandenen Zielordner an.'
        'Err_MissMapFile'        = 'Fehlende Zuordnungsdatei:'
        'Err_DupKey'             = 'Doppelter Schlüssel'
        'Err_InMap'              = 'in Zuordnung'
        'Err_MissMap'            = 'Fehlende Zuordnung'
        'Err_MissField'          = 'Fehlendes Feld'
        'Err_InvalidChars'       = 'Der Zielname enthält ungültige Zeichen.'
        'Err_DupDest'            = 'Doppelter Zielpfad'
        'Err_FileExists'         = 'Zieldatei existiert bereits'
        'Err_Blocked'            = 'Vorgang blockiert:'
        'Err_FixPrev'            = 'Fehler in der Vorschau. Beheben Sie diese vor dem Kopieren.'
        'Msg_CopyCount'          = 'Kopieren'
        'Msg_CopyOrig'           = 'Dateien? Originale werden nicht geändert.'
        'Msg_MoveOrig'           = 'Dateien? Originale werden verschoben.'
        'Log_Copied'             = 'Kopiert:'
        'Log_Moved'              = 'Verschoben:'
        'Log_CopyErr'            = 'Kopierfehler:'
        'Status_Copied'          = 'Abgeschlossen. Kopiert:'
        'Status_Moved'           = 'Abgeschlossen. Verschoben:'
        'Lbl_ProfileName'        = 'Profilname:'
        'Status_ProfSaved'       = 'Profil gespeichert:'
        'Status_ProfLoaded'      = 'Profil geladen:'
        'Status_AnalysisDone'    = 'Analyse abgeschlossen.'
        'Err_SelFieldTab'        = 'Wählen Sie ein Feld in der Tabelle.'
        'Err_FieldName'          = 'Feldname angeben.'
        'Err_SelFieldTrans'      = 'Wählen Sie ein Feld aus, dem Sie eine Transformation hinzufügen möchten.'
        'Err_SelFieldCombo'      = 'Wählen Sie ein Feld aus dem Dropdown-Menü.'
        'Tag_Field'              = '[Feld]'
        'Tag_Text'               = '[Text]'
        'Tag_Separator'          = '[Trennzeichen]'
        'Status_PrevBuilt'       = 'Vorschau erstellt.'
        'Err_BuildPrev1'         = 'Zuerst Vorschau erstellen.'
        'Status_Exported'        = 'Bericht exportiert:'
        'Err_SelProfList'        = 'Wählen Sie ein Profil aus der Liste.'
        'Err_SelProfCopy'        = 'Wählen Sie ein Profil zum Kopieren aus.'
        'Prefix_Copy'            = 'Kopie von '
        'Err_SelProfDel'         = 'Wählen Sie ein Profil zum Löschen aus.'
        'Msg_DelProf'            = 'Profil löschen'
        'Status_ProfDel'         = 'Profil gelöscht:'
        'Chk_EnforcePattern'     = 'Muster auf andere Dateien erzwingen'
        'Txt_TokenRegex'         = 'Regex-Muster:'
        'Tip_TokenRegex'         = 'Standard: (?&lt;value&gt;[^_\-\s]+)|(?&lt;sep&gt;[_\-\s]+)'
        'Tip_TokenRegexLabel'    = 'Der Regex muss jedes Segment des Dateinamens erfassen. Verwenden Sie (?&lt;sep&gt;...) für Trennzeichen; alle anderen Treffer sind Werte. Beispiel: (?&lt;value&gt;[^_\-\s]+)|(?&lt;sep&gt;[_\-\s]+).'
    }
}
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

#region XAML Definition
$xamlTemplate = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="{t:WinTitle}" Height="920" Width="1480"
        MinHeight="720" MinWidth="1120"
        WindowStartupLocation="CenterScreen"
        FontFamily="Segoe UI" FontSize="13" Background="#F5F7FA">
  <Window.Resources>
    <Style TargetType="Button">
      <Setter Property="Margin" Value="4"/>
      <Setter Property="Padding" Value="10,5"/>
      <Setter Property="MinHeight" Value="30"/>
    </Style>
    <Style TargetType="TextBox">
      <Setter Property="Margin" Value="4"/>
      <Setter Property="Padding" Value="5"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
    </Style>
    <Style TargetType="ComboBox">
      <Setter Property="Margin" Value="4"/>
      <Setter Property="MinHeight" Value="29"/>
    </Style>
    <Style TargetType="DataGrid">
      <Setter Property="Margin" Value="4"/>
      <Setter Property="IsReadOnly" Value="True"/>
      <Setter Property="CanUserAddRows" Value="False"/>
      <Setter Property="AutoGenerateColumns" Value="False"/>
      <Setter Property="AlternatingRowBackground" Value="#EDF2F7"/>
    </Style>
  </Window.Resources>

  <DockPanel>
    <!-- Pasek statusu -->
    <Border DockPanel.Dock="Bottom" Background="#E4EAF1" Padding="9">
      <DockPanel>
        <TextBlock x:Name="StatusText"/>
        <TextBlock x:Name="LogText" DockPanel.Dock="Right" Foreground="#52606D"/>
                <StackPanel DockPanel.Dock="Right" Orientation="Horizontal" VerticalAlignment="Center" Margin="0,0,15,0">
                    <Label Content="{t:Lbl_Language}" VerticalAlignment="Center" Padding="0" Margin="0,0,4,0"/>
                    <ComboBox x:Name="LanguageSelector" Width="100" MinHeight="22" Padding="4,2">
                        <ComboBoxItem Content="Polski" Tag="PL"/>
                        <ComboBoxItem Content="English" Tag="EN"/>
                        <ComboBoxItem Content="Deutsch" Tag="DE"/>
                    </ComboBox>
                </StackPanel>
      </DockPanel>
    </Border>

    <Grid Margin="8">
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="*"/>
      </Grid.RowDefinitions>

      <!-- Foldery -->
      <GroupBox Grid.Row="0" Header="{t:Folders_Header}">
        <Grid>
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="150"/>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="110"/>
          </Grid.ColumnDefinitions>
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
          </Grid.RowDefinitions>
          <Label Grid.Row="0" Content="{t:Source_Folder}"/>
          <TextBox x:Name="SourcePath" Grid.Row="0" Grid.Column="1"/>
          <Button x:Name="BrowseSource" Grid.Row="0" Grid.Column="2" Content="{t:Btn_Browse}"/>
          <Label Grid.Row="1" Content="{t:Dest_Folder}"/>
          <TextBox x:Name="DestinationPath" Grid.Row="1" Grid.Column="1"/>
          <Button x:Name="BrowseDestination" Grid.Row="1" Grid.Column="2" Content="{t:Btn_Browse}"/>
          <CheckBox x:Name="Recursive" Grid.Row="2" Grid.Column="1"
                    Content="{t:Chk_Recursive}"
                    Margin="4,2,4,4" VerticalAlignment="Center"/>
          <CheckBox x:Name="PreserveFolderStructure" Grid.Row="3" Grid.Column="1"
                    Content="{t:Chk_PreserveStructure}" IsChecked="True"
                    Margin="4,2,4,4" VerticalAlignment="Center"/>
        </Grid>
      </GroupBox>

      <!-- Zakładki -->
      <TabControl Grid.Row="1" Margin="4">

        <!-- ========== 1. Profil ========== -->
        <TabItem Header="{t:Tab_Profile}">
          <Grid Margin="12">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="380"/>
              <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <GroupBox Header="{t:Saved_Profiles}" Grid.Column="0">
              <DockPanel>
                <StackPanel DockPanel.Dock="Bottom">
                  <WrapPanel>
                    <Button x:Name="ProfileNew"    Content="{t:Btn_New}"/>
                    <Button x:Name="ProfileLoad"   Content="{t:Btn_Load}"/>
                    <Button x:Name="ProfileCopy"   Content="{t:Btn_Copy}"/>
                    <Button x:Name="ProfileDelete" Content="{t:Btn_Delete}"/>
                  </WrapPanel>
                  <Button x:Name="ProfileSave" Content="{t:Btn_Save_Config}" Background="#D6E9FC"/>
                </StackPanel>
                <ListBox x:Name="ProfileList" DisplayMemberPath="Name"/>
              </DockPanel>
            </GroupBox>
            <StackPanel Grid.Column="1" Margin="20,0,0,0">
              <TextBlock FontSize="18" FontWeight="SemiBold" Text="{t:Title_Wizard}"/>
              <TextBlock Margin="0,12,0,0" TextWrapping="Wrap"
                         Text="{t:Txt_ProfileHint}"/>
              <TextBlock Margin="0,18,0,0" FontWeight="SemiBold" Text="{t:Txt_ActiveProfile}"/>
              <TextBlock x:Name="CurrentProfile" Margin="0,4,0,0" Text="{t:Txt_Unsaved}"/>
            </StackPanel>
          </Grid>
        </TabItem>

        <!-- ========== 2. Analiza plików ========== -->
        <TabItem Header="{t:Tab_Analysis}">
          <Grid Margin="8">
            <Grid.RowDefinitions>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="*"/>
            </Grid.RowDefinitions>
            <WrapPanel>
              <Button x:Name="Analyze" Content="{t:Btn_Analyze}" Background="#D6E9FC"/>
              <TextBlock Text="{t:Txt_Extension}" Margin="14,9,0,0"/>
              <ComboBox x:Name="ExtensionFilter" Width="130"/>
              <TextBlock Text="{t:Txt_TokenRegex}" Margin="14,9,0,0" Foreground="#52606D" ToolTip="{t:Tip_TokenRegexLabel}"/>
              <TextBox x:Name="TokenRegex" Width="400" Text="(?&lt;value&gt;[^_\-\s]+)|(?&lt;sep&gt;[_\-\s]+)" ToolTip="{t:Tip_TokenRegex}" FontFamily="Consolas"/>
              <TextBlock x:Name="AnalysisInfo" Margin="14,9,0,0"/>
            </WrapPanel>
            <Grid Grid.Row="1">
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="*"/>
              </Grid.ColumnDefinitions>
              <GroupBox Grid.Column="0" Header="{t:Header_DetectedStructs}">
                <DataGrid x:Name="PatternGrid" SelectionMode="Single">
                  <DataGrid.Columns>
                    <DataGridTextColumn Header="{t:Col_Extension}" Binding="{Binding Extension}" Width="*"/>
                    <DataGridTextColumn Header="{t:Col_Structure}"    Binding="{Binding Display}"   Width="3*"/>
                    <DataGridTextColumn Header="{t:Col_Files}"        Binding="{Binding Count}"     Width="*"/>
                  </DataGrid.Columns>
                </DataGrid>
              </GroupBox>
              <GroupBox Grid.Column="1" Header="{t:Header_Examples}">
                <DockPanel>
                  <TextBlock x:Name="PatternHint" DockPanel.Dock="Bottom" Margin="4"
                             TextWrapping="Wrap" Foreground="#52606D"/>
                  <ListBox x:Name="PatternSamples"/>
                </DockPanel>
              </GroupBox>
            </Grid>
          </Grid>
        </TabItem>

        <!-- ========== 3. Pola i mapowania ========== -->
        <TabItem Header="{t:Tab_Fields}">
          <Grid Margin="8">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>

            <!-- Lewa: tabela pól -->
            <GroupBox Grid.Column="0" Header="{t:Header_DetectedFields}">
              <DockPanel>
                <TextBlock DockPanel.Dock="Bottom" Margin="4" Foreground="#52606D" TextWrapping="Wrap"
                           Text="{t:Txt_FieldHint}"/>
                <DataGrid x:Name="FieldGrid" SelectionMode="Single">
                  <DataGrid.Columns>
                    <DataGridTextColumn Header="#"        Binding="{Binding DisplayIndex}" Width="30"/>
                    <DataGridTextColumn Header="{t:Col_Sample}" Binding="{Binding Sample}"       Width="*"/>
                    <DataGridTextColumn Header="{t:Col_Type}"      Binding="{Binding EffectiveType}" Width="*"/>
                    <DataGridTextColumn Header="{t:Col_Name}"    Binding="{Binding Name}"         Width="*"/>
                    <DataGridTextColumn Header="{t:Col_Role}"     Binding="{Binding Role}"         Width="*"/>
                    <DataGridTextColumn Header="{t:Col_Source}"   Binding="{Binding Source}"       Width="72"/>
                  </DataGrid.Columns>
                </DataGrid>
              </DockPanel>
            </GroupBox>

            <!-- Prawa: edycja pola -->
            <GroupBox Grid.Column="1" Header="{t:Header_EditField}">
              <ScrollViewer VerticalScrollBarVisibility="Auto">
                <StackPanel Margin="4">
                  <TextBlock Text="{t:Txt_FieldName}" Margin="0,0,0,2"/>
                  <TextBox x:Name="FieldName"/>
                  <TextBlock Text="{t:Txt_Role}" Margin="0,4,0,2"/>
                  <ComboBox x:Name="FieldRole">
                    <ComboBoxItem Content="{t:Role_Value}"/>
                    <ComboBoxItem Content="{t:Role_Date}"/>
                    <ComboBoxItem Content="{t:Role_Id}"/>
                    <ComboBoxItem Content="{t:Role_Const}"/>
                    <ComboBoxItem Content="{t:Role_Ignore}"/>
                  </ComboBox>
                                    <TextBlock Text="{t:Txt_DataType}" Margin="0,4,0,2"/>
                                    <ComboBox x:Name="FieldType" DisplayMemberPath="Label"/>
                  <Button x:Name="FieldApply" Content="{t:Btn_ApplyField}" Margin="4,6,4,4"/>

                  <Separator Margin="4,8"/>

                  <!-- Transformacje per-pole -->
                  <TextBlock FontWeight="Bold" Text="{t:Header_Transformations}" Margin="0,4,0,4"/>
                  <TextBlock Foreground="#52606D" TextWrapping="Wrap" FontSize="11.5"
                             Text="{t:Txt_TransformHint}"/>
                  <ListBox x:Name="TransformList" DisplayMemberPath="Display" Height="105" Margin="4"/>
                  <WrapPanel>
                    <Button x:Name="TransformAdd"    Content="{t:Btn_AddTransform}"/>
                    <Button x:Name="TransformRemove" Content="{t:Btn_Delete}"/>
                  </WrapPanel>

                  <Separator Margin="4,8"/>

                  <!-- Mapowania globalne -->
                  <TextBlock FontWeight="Bold" Text="{t:Header_Mappings}" Margin="0,4,0,4"/>
                  <TextBlock Foreground="#52606D" TextWrapping="Wrap" FontSize="11.5"
                             Text="{t:Txt_MappingHint}"/>
                  <ListBox x:Name="MappingList" DisplayMemberPath="Display" Height="105" Margin="4"/>
                  <WrapPanel>
                    <Button x:Name="MappingAdd"    Content="{t:Btn_AddMapping}"/>
                    <Button x:Name="MappingEdit"   Content="{t:Btn_Edit}"/>
                    <Button x:Name="MappingRemove" Content="{t:Btn_Delete}"/>
                    <Button x:Name="MappingUp"     Content="{t:Btn_Up}"/>
                    <Button x:Name="MappingDown"   Content="{t:Btn_Down}"/>
                  </WrapPanel>
                </StackPanel>
              </ScrollViewer>
            </GroupBox>
          </Grid>
        </TabItem>

        <!-- ========== 4. Nazwa docelowa ========== -->
        <TabItem Header="{t:Tab_DestName}">
          <Grid Margin="8">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>

            <!-- Lewa: budowanie nazwy -->
            <GroupBox Grid.Column="0" Header="{t:Header_DestElements}">
              <DockPanel>
                <StackPanel DockPanel.Dock="Bottom">
                  <WrapPanel Margin="0,4,0,0">
                    <ComboBox x:Name="FieldSelector" Width="240"/>
                    <Button x:Name="OutputAddField" Content="{t:Btn_AddField}"/>
                  </WrapPanel>
                  <WrapPanel>
                    <TextBox x:Name="OutputText" Width="240"/>
                    <Button x:Name="OutputAddText"      Content="{t:Btn_AddText}"/>
                    <Button x:Name="OutputAddSeparator" Content="{t:Btn_AddSep}"/>
                  </WrapPanel>
                  <WrapPanel>
                    <Button x:Name="OutputUp"     Content="{t:Btn_Up}"/>
                    <Button x:Name="OutputDown"   Content="{t:Btn_Down}"/>
                    <Button x:Name="OutputRemove" Content="{t:Btn_RemoveElement}"/>
                  </WrapPanel>
                </StackPanel>
                <ListBox x:Name="OutputList" DisplayMemberPath="Display"/>
              </DockPanel>
            </GroupBox>

            <!-- Prawa: ustawienia rozszerzenia i podgląd -->
            <GroupBox Grid.Column="1" Header="{t:Header_Settings}">
              <StackPanel Margin="4">
                <WrapPanel>
                  <CheckBox x:Name="KeepExtension" IsChecked="True"
                            Content="{t:Chk_KeepExt}" Margin="4"/>
                </WrapPanel>
                <WrapPanel>
                  <TextBlock Text="{t:Txt_NewExt}" Margin="10,8,4,0"/>
                  <TextBox x:Name="NewExtension" Width="120" IsEnabled="False"/>
                </WrapPanel>

                <Separator Margin="4,12"/>

                <TextBlock FontWeight="Bold" Text="{t:Txt_PreviewTitle}" Margin="0,4,0,8"/>
                <Border Background="#FFFEF5" BorderBrush="#E0D8B0" BorderThickness="1"
                        Padding="10" CornerRadius="4">
                  <TextBlock x:Name="OutputExample" TextWrapping="Wrap"
                             FontFamily="Consolas" FontSize="13"/>
                </Border>
              </StackPanel>
            </GroupBox>
          </Grid>
        </TabItem>

        <!-- ========== 5. Podgląd i wykonanie ========== -->
        <TabItem Header="{t:Tab_Preview}">
          <DockPanel Margin="8">
            <StackPanel DockPanel.Dock="Top" Orientation="Horizontal">
              <Button x:Name="BuildPreview" Content="{t:Btn_BuildPreview}" Background="#D6E9FC"/>
              <CheckBox x:Name="EnforcePattern" Content="{t:Chk_EnforcePattern}" Margin="10,0" VerticalAlignment="Center"/>
              <ComboBox x:Name="PreviewFilter" Width="150">
                <ComboBoxItem Content="{t:Cbo_All}" IsSelected="True"/>
                <ComboBoxItem Content="{t:Cbo_Errors}"/>
                <ComboBoxItem Content="{t:Cbo_Ready}"/>
              </ComboBox>
              <TextBlock x:Name="PreviewInfo" Margin="12,10,4,4"/>
            </StackPanel>
            <StackPanel DockPanel.Dock="Bottom">
              <ProgressBar x:Name="ProgressBar" Height="18" Margin="4" Visibility="Collapsed"/>
              <WrapPanel HorizontalAlignment="Right">
                <TextBlock Text="{t:Lbl_ExecutionMode}" Margin="0,0,6,0" VerticalAlignment="Center"/>
                <ComboBox x:Name="ExecutionMode" Width="110" Margin="0,0,8,0">
                  <ComboBoxItem Content="{t:Action_Copy}" IsSelected="True"/>
                  <ComboBoxItem Content="{t:Action_Move}"/>
                </ComboBox>
                <Button x:Name="ExportAudit" Content="{t:Btn_Export}"/>
                <Button x:Name="OpenLog"     Content="{t:Btn_OpenLog}"/>
                <Button x:Name="Execute"     Content="{t:Btn_Execute}" Background="#C6EFCE"/>
              </WrapPanel>
            </StackPanel>
            <DataGrid x:Name="PreviewGrid">
              <DataGrid.RowStyle>
                <Style TargetType="DataGridRow">
                  <Setter Property="ToolTip" Value="{Binding Details}"/>
                  <Style.Triggers>
                    <DataTrigger Binding="{Binding StatusCode}" Value="Error">
                      <Setter Property="Background" Value="#FFE0E0"/>
                    </DataTrigger>
                    <DataTrigger Binding="{Binding StatusCode}" Value="Ready">
                      <Setter Property="Background" Value="#E8F5E9"/>
                    </DataTrigger>
                  </Style.Triggers>
                </Style>
              </DataGrid.RowStyle>
              <DataGrid.Columns>
                <DataGridTextColumn Header="{t:Col_SourceFile}"  Binding="{Binding SourceRelative}"      Width="2*"/>
                <DataGridTextColumn Header="{t:Col_DestName}" Binding="{Binding DestinationRelative}"  Width="2*"/>
                <DataGridTextColumn Header="{t:Col_Status}"         Binding="{Binding Status}"               Width="80"/>
                <DataGridTextColumn Header="{t:Col_Details}"      Binding="{Binding Details}"              Width="3*"/>
              </DataGrid.Columns>
            </DataGrid>
          </DockPanel>
        </TabItem>

      </TabControl>
    </Grid>
  </DockPanel>
</Window>
'@
#endregion

#region Window Creation and Element Binding
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

function ErrorBox([string]$title, $err) {
    $msg = "Komunikat: $($err.Exception.Message)"
    $msg += "`nLinia:     $($err.InvocationInfo.ScriptLineNumber)"
    $msg += "`nLog:       $script:LogPath"
    Log "$title | $($err.Exception.Message)" 'ERROR'
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

#region Tokenization and Analysis

function Tokens([string]$name) {
    $pattern = if ($TokenRegex -and $TokenRegex.Text) { $TokenRegex.Text } else { '(?<value>[^_\-\s]+)|(?<sep>[_\-\s]+)' }
    Get-FNTTokens -Name $name -Pattern $pattern -CustomTypeRules @($script:Config.CustomTypeRules)
}

function TokenTypeLabel([string]$typeId) {
    switch ($typeId) {
        'Integer' { return (T 'Type_Integer') }
        'Decimal' { return (T 'Type_Decimal') }
        'DateTime' { return (T 'Type_DateTime') }
        'Guid' { return (T 'Type_Guid') }
        'Version' { return (T 'Type_Version') }
        'Ambiguous' { return (T 'Type_Ambiguous') }
        default { return (T 'Type_Text') }
    }
}

function ConvertLegacyTypeId([string]$typeLabel) {
    if ([string]::IsNullOrWhiteSpace($typeLabel)) { return 'Text' }
    if ($typeLabel -match 'Ambig|Niejedno|Mehrdeu') { return 'Ambiguous' }
    if ($typeLabel -match 'GUID') { return 'Guid' }
    if ($typeLabel -match 'Version|Wersja') { return 'Version' }
    if ($typeLabel -match 'Decimal|dzies|Dezimal') { return 'Decimal' }
    if ($typeLabel -match 'Date|Data|Datum') { return 'DateTime' }
    if ($typeLabel -match 'Number|Liczba|Zahl') { return 'Integer' }
    return 'Text'
}

function GetFieldTypeOptions($field) {
    $options = @(
        [pscustomobject]@{ Id = 'Auto'; Format = $null; Label = (T 'Type_Auto') }
        [pscustomobject]@{ Id = 'Text'; Format = $null; Label = (T 'Type_Text') }
        [pscustomobject]@{ Id = 'Integer'; Format = $null; Label = (T 'Type_Integer') }
        [pscustomobject]@{ Id = 'Decimal'; Format = $null; Label = (T 'Type_Decimal') }
        [pscustomobject]@{ Id = 'Guid'; Format = $null; Label = (T 'Type_Guid') }
        [pscustomobject]@{ Id = 'Version'; Format = $null; Label = (T 'Type_Version') }
    )

    foreach ($candidate in @($field.CandidateTypes)) {
        if ($candidate.TypeId -eq 'DateTime') {
            $label = "$(T 'Type_DateTime') ($($candidate.Format))"
            $options += [pscustomobject]@{ Id = 'DateTime'; Format = [string]$candidate.Format; Label = $label }
        }
        elseif ([string]$candidate.TypeId -like 'Custom:*') {
            $options += [pscustomobject]@{
                Id     = [string]$candidate.TypeId
                Format = $null
                Label  = [string]$candidate.TypeId
            }
        }
    }

    $seen = @{}
    @($options | Where-Object {
            $key = "$($_.Id)|$($_.Format)"
            if ($seen.ContainsKey($key)) { return $false }
            $seen[$key] = $true
            return $true
        })
}

function GetEffectiveTypeLabel($field) {
    if (-not $field.SelectedTypeId -or $field.SelectedTypeId -eq 'Auto') {
        return (TokenTypeLabel ([string]$field.DetectedTypeId))
    }
    if ($field.SelectedTypeId -eq 'DateTime' -and $field.SelectedFormat) {
        return "$(T 'Type_DateTime') ($($field.SelectedFormat))"
    }
    if ([string]$field.SelectedTypeId -like 'Custom:*') {
        return [string]$field.SelectedTypeId
    }
    return (TokenTypeLabel ([string]$field.SelectedTypeId))
}

function GetResolvedFieldType($field) {
    if ($field.SelectedTypeId -and $field.SelectedTypeId -ne 'Auto') {
        return [pscustomobject]@{
            TypeId = [string]$field.SelectedTypeId
            Format = [string]$field.SelectedFormat
        }
    }

    $format = $null
    if ($field.DetectedTypeId -eq 'DateTime') {
        $dateCandidates = @($field.CandidateTypes | Where-Object { $_.TypeId -eq 'DateTime' })
        if ($dateCandidates.Count -eq 1) {
            $format = [string]$dateCandidates[0].Format
        }
    }
    return [pscustomobject]@{
        TypeId = [string]$field.DetectedTypeId
        Format = $format
    }
}

function ValidateFieldValues([hashtable]$values) {
    foreach ($field in $script:Fields) {
        if ($field.IsVirtual -or -not $values.ContainsKey($field.Name)) { continue }
        $resolvedType = GetResolvedFieldType $field
        if (-not (Test-FNTValueType -Value ([string]$values[$field.Name]) -TypeId $resolvedType.TypeId `
                -Format $resolvedType.Format -CustomTypeRules @($script:Config.CustomTypeRules))) {
            $formatText = if ($resolvedType.Format) { " ($($resolvedType.Format))" } else { '' }
            throw "$($field.Name): '$($values[$field.Name])' != $($resolvedType.TypeId)$formatText"
        }
    }
}

function ParseNameByTemplate([string]$name, $templateParts) {
    $fieldTypes = @{}
    foreach ($field in $script:Fields) {
        if (-not $field.IsVirtual -and $field.PartIndex -ge 0) {
            $fieldTypes[[int]$field.PartIndex] = GetResolvedFieldType $field
        }
    }
    $pattern = if ($TokenRegex -and $TokenRegex.Text) { $TokenRegex.Text } else { '(?<value>[^_\-\s]+)|(?<sep>[_\-\s]+)' }
    $match = Match-FNTNamePattern -Name $name -PatternTokens @($templateParts) -TokenizerPattern $pattern `
        -FieldTypes $fieldTypes -CustomTypeRules @($script:Config.CustomTypeRules)
    return $match.Values
}

function AnalyzePatterns {
    $src = $SourcePath.Text.Trim()
    if (-not (Test-Path $src -PathType Container)) {
        throw (T 'Err_SrcNotExist')
    }

    $files = if ($Recursive.IsChecked) {
        @(Get-ChildItem -LiteralPath $src -File -Recurse)
    }
    else {
        @(Get-ChildItem -LiteralPath $src -File)
    }
    if (-not $files) { throw (T 'Err_NoFiles') }

    # Populate extension filter
    $extensions = @($files | ForEach-Object { $_.Extension.ToLower() } | Sort-Object -Unique)
    $ExtensionFilter.ItemsSource = $extensions
    if (-not $ExtensionFilter.SelectedItem) {
        $ExtensionFilter.SelectedItem = ($extensions | Select-Object -First 1)
    }

    BuildPatternList
}

function BuildPatternList {
    $src = $SourcePath.Text.Trim()
    $ext = [string]$ExtensionFilter.SelectedItem
    if (-not $ext) { return }

    $files = if ($Recursive.IsChecked) {
        @(Get-ChildItem -LiteralPath $src -File -Recurse -Filter "*$ext")
    }
    else {
        @(Get-ChildItem -LiteralPath $src -File -Filter "*$ext")
    }

    # Group files by token structure signature
    $raw = @()
    foreach ($f in $files) {
        $parts = Tokens $f.BaseName
        $sig = ($parts | ForEach-Object {
                if ($_.IsSeparator) { 'S:' + $_.Value + '|' }
            else { 'T:' + $_.DetectedTypeId + '|' }
            }) -join ''
        $raw += [pscustomobject]@{ Signature = $sig; Parts = $parts; File = $f }
    }

    $script:Patterns = @()
    foreach ($g in $raw | Group-Object Signature) {
        $sample = $g.Group[0]
        $all = $g.Group
        $labels = @()
        $fieldInferences = @{}
        for ($i = 0; $i -lt $sample.Parts.Count; $i++) {
            if ($sample.Parts[$i].IsSeparator) {
                $labels += $sample.Parts[$i].Value
                continue
            }
            $uniqueValues = @($all | ForEach-Object { $_.Parts[$i].Value } | Select-Object -Unique)
            $inference = Get-FNTFieldInference -Tokens @($all | ForEach-Object { $_.Parts[$i] })
            $fieldInferences[$i] = $inference
            $kind = TokenTypeLabel $inference.DetectedTypeId
            $labels += if ($uniqueValues.Count -eq 1) { "[$($uniqueValues[0])]" } else { "<$kind>" }
        }
        $script:Patterns += [pscustomobject]@{
            Extension = $ext
            Display   = ($labels -join '')
            Count     = $g.Count
            Items     = $all
            Signature = $g.Name
            FieldInferences = $fieldInferences
        }
    }

    $PatternGrid.ItemsSource = $script:Patterns
    $AnalysisInfo.Text = "$(T 'Msg_Files'): $($files.Count); $(T 'Msg_Structs'): $($script:Patterns.Count)"
}

function SetPattern($pattern) {
    $script:CurrentPattern = $pattern

    # Show sample file names
    $PatternSamples.ItemsSource = @(
        $pattern.Items | Select-Object -First 20 | ForEach-Object { $_.File.Name }
    )

    # Build fields from token positions (clear old fields)
    $script:Fields.Clear()
    $parts = $pattern.Items[0].Parts
    for ($i = 0; $i -lt $parts.Count; $i++) {
        if ($parts[$i].IsSeparator) { continue }

        $uniqueValues = @($pattern.Items | ForEach-Object { $_.Parts[$i].Value } | Select-Object -Unique)
        $inference = $pattern.FieldInferences[$i]
        $typeId = $inference.DetectedTypeId
        $type = TokenTypeLabel $typeId

        # Auto-detect role
        $role = if ($typeId -eq 'DateTime') { (T 'Role_Date') }
        elseif ($uniqueValues.Count -eq 1) { (T 'Role_Const') }
        else { (T 'Role_Value') }

        # Auto-name
        $name = switch ($role) {
            (T 'Role_Date') { "$(T 'Name_Date')_$($i + 1)" }
            (T 'Role_Const') { "$(T 'Name_Text')_$($i + 1)" }
            default { "$(T 'Name_Field')_$($i + 1)" }
        }

        $script:Fields.Add([pscustomobject]@{
                PartIndex    = $i
                DisplayIndex = "$($script:Fields.Count + 1)"
                Sample       = $parts[$i].Value
                DetectedType = $type
                DetectedTypeId = $typeId
                CandidateTypes = @($inference.CandidateTypes)
                IsAmbiguous  = [bool]$inference.IsAmbiguous
                SelectedTypeId = 'Auto'
                SelectedFormat = $null
                EffectiveType = $type
                Name         = $name
                Role         = $role
                IsVirtual    = $false
                Source       = (T 'Src_Name')
                Transforms   = [System.Collections.ArrayList]::new()
            })
    }

    # Recreate virtual fields from existing mappings
    foreach ($m in $script:Mappings) {
        EnsureVirtualField $m.OutputField
    }

    $FieldGrid.ItemsSource = $script:Fields
    RefreshFieldSelector
    $PatternHint.Text = (T 'Hint_SelectField')
    UpdateOutputExample
}

#endregion

#region Transform Functions

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
            }
        }
        catch {
            throw "$(T 'Err_Transform') '$($t.Display)': $($_.Exception.Message)"
        }
    }
    return $value
}

function ShowTransformDialog {
    $form = New-Object Windows.Forms.Form
    $form.Text = (T 'Title_AddTransform')
    $form.Size = New-Object Drawing.Size(520, 300)
    $form.StartPosition = 'CenterParent'
    $form.Font = New-Object Drawing.Font('Segoe UI', 9)
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false

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
                    $paramControls[0].Text = 'yyyyMMdd'; $paramControls[0].Visible = $true
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

#endregion

#region Mapping Functions

function CsvHeaders([string]$path) {
    $headerLine = Get-Content -LiteralPath $path -Encoding UTF8 | Where-Object { $_ -and $_.Trim() } | Select-Object -First 1
    if (-not $headerLine) {
        return [pscustomobject]@{ Delimiter = ','; Headers = @() }
    }

    # Auto-detect delimiter by frequency
    $delimiter = @(';', ',', "`t", '|') |
    Sort-Object { ([regex]::Matches($headerLine, [regex]::Escape($_))).Count } -Descending |
    Select-Object -First 1

    $headers = @()
    if ($headerLine) {
        $headers = @($headerLine -split [regex]::Escape($delimiter))
        $headers = @($headers | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' })
    }

    [pscustomobject]@{ Delimiter = $delimiter; Headers = $headers }
}

function EnsureVirtualField([string]$name) {
    $exists = $false
    foreach ($f in $script:Fields) {
        if ($f.Name -eq $name) { $exists = $true; break }
    }
    if (-not $exists) {
        $script:Fields.Add([pscustomobject]@{
                PartIndex    = -1
                DisplayIndex = 'V'
                Sample       = (T 'Val_Mapping')
                DetectedType = (T 'Src_Mapping')
                Name         = $name
                Role         = (T 'Role_Value')
                IsVirtual    = $true
                Source       = (T 'Src_Mapping')
                Transforms   = [System.Collections.ArrayList]::new()
            })
    }
}

function AddMappingDialog([object]$mapping = $null) {
    $form = New-Object Windows.Forms.Form
    $form.Text = if ($mapping) { (T 'Title_EditMapping') } else { (T 'Title_AddMapping') }
    $form.Size = New-Object Drawing.Size(650, 310)
    $form.StartPosition = 'CenterParent'
    $form.Font = New-Object Drawing.Font('Segoe UI', 9)
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false

    $labels = @(
        (T 'Lbl_MapName'),
        (T 'Lbl_MapIn'),
        (T 'Lbl_MapOut'),
        (T 'Lbl_MapFile'),
        (T 'Lbl_MapKey'),
        (T 'Lbl_MapVal')
    )
    $controls = @()

    for ($i = 0; $i -lt 6; $i++) {
        $lbl = New-Object Windows.Forms.Label
        $lbl.Text = $labels[$i]
        $lbl.Location = New-Object Drawing.Point(15, (15 + $i * 35))
        $lbl.Size = New-Object Drawing.Size(140, 25)
        $form.Controls.Add($lbl)

        if ($i -in 1, 4, 5) {
            $ctl = New-Object Windows.Forms.ComboBox
            $ctl.DropDownStyle = 'DropDownList'
        }
        else {
            $ctl = New-Object Windows.Forms.TextBox
        }
        $ctl.Location = New-Object Drawing.Point(160, (12 + $i * 35))
        $ctl.Size = New-Object Drawing.Size(350, 25)
        $form.Controls.Add($ctl)
        $controls += $ctl
    }

    # Populate input field combo with all field names
    $controls[1].Items.AddRange(@($script:Fields | ForEach-Object { $_.Name }))

    if ($mapping) {
        $controls[0].Text = [string]$mapping.Name
        $controls[1].SelectedItem = $mapping.InputField
        $controls[2].Text = [string]$mapping.OutputField
        $controls[3].Text = [string]$mapping.Path
        $controls[4].SelectedItem = $mapping.KeyColumn
        $controls[5].SelectedItem = $mapping.ValueColumn
        if ($mapping.Path) {
            try {
                $h = CsvHeaders $mapping.Path
                $script:tempDelimiter = $h.Delimiter
                $controls[4].Items.Clear()
                $controls[5].Items.Clear()
                $controls[4].Items.AddRange($h.Headers)
                $controls[5].Items.AddRange($h.Headers)
            }
            catch {
                $script:tempDelimiter = $mapping.Delimiter
            }
        }
        else {
            $script:tempDelimiter = $mapping.Delimiter
        }
    }

    $script:tempDelimiter = if ($mapping -and $mapping.Delimiter) { [string]$mapping.Delimiter } else { ',' }

    # Browse button for data file
    $browse = New-Object Windows.Forms.Button
    $browse.Text = $(T 'Btn_Browse')
    $browse.Location = New-Object Drawing.Point(520, 117)
    $browse.Size = New-Object Drawing.Size(80, 25)
    $form.Controls.Add($browse)

    $browse.Add_Click({
            $p = FileDialog
            if ($p) {
                $controls[3].Text = $p
                try {
                    $h = CsvHeaders $p
                    $script:tempDelimiter = $h.Delimiter
                    $controls[4].Items.Clear()
                    $controls[5].Items.Clear()
                    $controls[4].Items.AddRange($h.Headers)
                    $controls[5].Items.AddRange($h.Headers)
                }
                catch {
                    [Windows.Forms.MessageBox]::Show(
                        "$(T 'Err_ReadHeaders') $($_.Exception.Message)",
                        (T 'Title_FileErr'), 'OK', 'Warning'
                    ) | Out-Null
                }
            }
        })

    # Add button
    $btnOk = New-Object Windows.Forms.Button
    $btnOk.Text = if ($mapping) { (T 'Btn_Edit') } else { (T 'Btn_Add') }
    $btnOk.Location = New-Object Drawing.Point(520, 225)
    $btnOk.Size = New-Object Drawing.Size(80, 30)
    $form.Controls.Add($btnOk)

    $btnOk.Add_Click({
            if (-not $controls[0].Text -or
                $controls[1].SelectedItem -eq $null -or
                -not $controls[2].Text -or
                -not $controls[3].Text -or
                $controls[4].SelectedItem -eq $null -or
                $controls[5].SelectedItem -eq $null) {
                [Windows.Forms.MessageBox]::Show((T 'Err_FillFields'), (T 'Title_MissingData'), 'OK', 'Warning') | Out-Null
                return
            }

            $targetMapping = $mapping
            if ($targetMapping) {
                $oldOutputField = [string]$targetMapping.OutputField
                $targetMapping.Name = $controls[0].Text
                $targetMapping.InputField = [string]$controls[1].SelectedItem
                $targetMapping.OutputField = $controls[2].Text.Trim()
                $targetMapping.Path = $controls[3].Text
                $targetMapping.KeyColumn = [string]$controls[4].SelectedItem
                $targetMapping.ValueColumn = [string]$controls[5].SelectedItem
                $targetMapping.Delimiter = $script:tempDelimiter
                $targetMapping.Display = "$($controls[0].Text): $($controls[1].SelectedItem) → $($controls[2].Text)"
            }
            else {
                $targetMapping = [pscustomobject]@{
                    Name        = $controls[0].Text
                    InputField  = [string]$controls[1].SelectedItem
                    OutputField = $controls[2].Text.Trim()
                    Path        = $controls[3].Text
                    KeyColumn   = [string]$controls[4].SelectedItem
                    ValueColumn = [string]$controls[5].SelectedItem
                    Delimiter   = $script:tempDelimiter
                    Display     = "$($controls[0].Text): $($controls[1].SelectedItem) → $($controls[2].Text)"
                }
                $script:Mappings.Add($targetMapping)
            }

            # Create or refresh virtual field for the output
            $newOutputField = [string]$targetMapping.OutputField
            if ($mapping) {
                $oldOutputField = [string]$mapping.OutputField
                if ($oldOutputField -and $oldOutputField -ne $newOutputField) {
                    $stillUsed = $false
                    foreach ($m in $script:Mappings) {
                        if ($m.OutputField -eq $oldOutputField) { $stillUsed = $true; break }
                    }
                    if (-not $stillUsed) {
                        $toRemove = $null
                        foreach ($f in $script:Fields) {
                            if ($f.IsVirtual -and $f.Name -eq $oldOutputField) { $toRemove = $f; break }
                        }
                        if ($toRemove) { $script:Fields.Remove($toRemove) }
                    }
                }
            }
            EnsureVirtualField $newOutputField

            # Refresh UI
            $MappingList.ItemsSource = $null
            $MappingList.ItemsSource = $script:Mappings
            $FieldGrid.ItemsSource = $null
            $FieldGrid.ItemsSource = $script:Fields
            RefreshFieldSelector

            $form.Close()
        })

    [void]$form.ShowDialog()
}

#endregion

#region Output, Preview, and Execution

function RefreshFieldSelector {
    $names = @($script:Fields | ForEach-Object { $_.Name })
    $prev = $FieldSelector.SelectedItem
    $FieldSelector.ItemsSource = $names
    if ($prev -and $names -contains $prev) {
        $FieldSelector.SelectedItem = $prev
    }
}

function UpdateOutputExample {
    if (-not $script:OutputParts.Count) {
        $OutputExample.Text = (T 'Hint_AddElements')
        return
    }

    # Structure display
    $structParts = @($script:OutputParts | ForEach-Object { $_.Display })
    $structure = ($structParts -join '  +  ')

    # Try live preview from first file
    $preview = ''
    if ($script:CurrentPattern -and $script:CurrentPattern.Items.Count -gt 0) {
        try {
            $item = $script:CurrentPattern.Items[0]
            $values = @{}

            # 1. Extract raw values from filename fields
            foreach ($f in $script:Fields) {
                if (-not $f.IsVirtual -and $f.PartIndex -ge 0) {
                    $values[$f.Name] = $item.Parts[$f.PartIndex].Value
                }
            }

            # 2. Apply mappings (in order, allows chaining)
            foreach ($m in $script:Mappings) {
                if (-not (Test-Path $m.Path)) { continue }
                $csv = Import-Csv -LiteralPath $m.Path -Delimiter $m.Delimiter
                $inputVal = [string]$values[$m.InputField]
                $match = $csv | Where-Object {
                    ([string]$_.($m.KeyColumn)).Trim() -eq $inputVal
                } | Select-Object -First 1
                if ($match) {
                    $values[$m.OutputField] = ([string]$match.($m.ValueColumn)).Trim()
                }
            }

            # 3. Apply per-field transforms
            foreach ($f in $script:Fields) {
                if ($values.ContainsKey($f.Name) -and $f.Transforms -and $f.Transforms.Count -gt 0) {
                    $values[$f.Name] = ApplyTransforms $values[$f.Name] $f.Transforms
                }
            }

            # 4. Build output name
            $name = ''
            foreach ($p in $script:OutputParts) {
                if ($p.Type -eq 'Text') {
                    $name += $p.Value
                }
                elseif ($values.ContainsKey($p.Value)) {
                    $name += $values[$p.Value]
                }
                else {
                    $name += "?$($p.Value)?"
                }
            }

            $ext = if ($KeepExtension.IsChecked) { $item.File.Extension }
            else { $NewExtension.Text.Trim() }
            if ($ext -and -not $ext.StartsWith('.')) { $ext = '.' + $ext }

            $preview = "`n`n$(T 'Prefix_Source'):   $($item.File.Name)`n$(T 'Result'):   $name$ext"
        }
        catch {
            $preview = "`n`n($(T 'PreviewUnavailable'): $($_.Exception.Message))"
        }
    }

    $OutputExample.Text = "$structure$preview"
}

function FullBuildPreview {
    $script:PreviewRows.Clear()

    if (-not $script:CurrentPattern) {
        throw (T 'Err_SelectPattern')
    }
    if (-not $script:OutputParts.Count) {
        throw (T 'Err_AddDest')
    }
    foreach ($field in $script:Fields) {
        if (-not $field.IsVirtual -and $field.IsAmbiguous -and $field.SelectedTypeId -eq 'Auto') {
            throw "$(T 'Err_ResolveAmbiguous') $($field.Name)"
        }
    }

    $src = $SourcePath.Text.Trim()
    $dst = $DestinationPath.Text.Trim()
    if (-not (Test-Path $dst -PathType Container)) {
        throw (T 'Err_DestNotExist')
    }

    # Pre-load all mapping files into hashtables
    $maps = @{}
    foreach ($def in $script:Mappings) {
        if (-not (Test-Path $def.Path)) {
            throw "$(T 'Err_MissMapFile') $($def.Path)"
        }
        $h = @{}
        Import-Csv -LiteralPath $def.Path -Delimiter $def.Delimiter | ForEach-Object {
            $k = ([string]$_.($def.KeyColumn)).Trim()
            if ($h.ContainsKey($k)) {
                throw "$(T 'Err_DupKey') '$k' $(T 'Err_InMap') '$($def.Name)'."
            }
            $h[$k] = ([string]$_.($def.ValueColumn)).Trim()
        }
        $maps[$def.Name] = $h
    }

    # Process each file
    $filesToProcess = if ($EnforcePattern.IsChecked) {
        $all = @()
        foreach ($p in $script:Patterns) { $all += $p.Items }
        $all
    }
    else {
        $script:CurrentPattern.Items
    }

    foreach ($item in $filesToProcess) {
        $row = [ordered]@{
            SourcePath          = $item.File.FullName
            SourceRelative      = $item.File.FullName.Substring($src.TrimEnd('\').Length).TrimStart('\')
            DestinationPath     = ''
            DestinationRelative = ''
            StatusCode          = 'Ready'
            Status              = (T 'Cbo_Ready')
            Details             = ''
        }

        try {
            $values = @{}

            # 1. Extract raw values
            if ($EnforcePattern.IsChecked) {
                $parsedDict = ParseNameByTemplate $item.File.BaseName $script:CurrentPattern.Items[0].Parts
                foreach ($f in $script:Fields) {
                    if (-not $f.IsVirtual -and $f.PartIndex -ge 0) {
                        $values[$f.Name] = $parsedDict[$f.PartIndex]
                    }
                }
            }
            else {
                foreach ($f in $script:Fields) {
                    if (-not $f.IsVirtual -and $f.PartIndex -ge 0) {
                        $values[$f.Name] = $item.Parts[$f.PartIndex].Value
                    }
                }
            }
            ValidateFieldValues $values

            # 2. Apply mappings (in order)
            foreach ($m in $script:Mappings) {
                $key = [string]$values[$m.InputField]
                if (-not $maps[$m.Name].ContainsKey($key)) {
                    throw "$(T 'Err_MissMap') '$key' $(T 'Err_InMap') '$($m.Name)'."
                }
                $values[$m.OutputField] = $maps[$m.Name][$key]
            }

            # 3. Apply per-field transforms
            foreach ($f in $script:Fields) {
                if ($values.ContainsKey($f.Name) -and $f.Transforms -and $f.Transforms.Count -gt 0) {
                    $values[$f.Name] = ApplyTransforms $values[$f.Name] $f.Transforms
                }
            }

            # 4. Build output name
            $name = ''
            foreach ($p in $script:OutputParts) {
                if ($p.Type -eq 'Text') {
                    $name += $p.Value
                }
                else {
                    if (-not $values.ContainsKey($p.Value)) {
                        throw "$(T 'Err_MissField') '$($p.Value)'."
                    }
                    $name += $values[$p.Value]
                }
            }

            # Validate filename characters
            if ($name.IndexOfAny([IO.Path]::GetInvalidFileNameChars()) -ge 0) {
                throw ($(T 'Err_InvalidChars') + " $name")
            }

            # Extension
            $ext = if ($KeepExtension.IsChecked) { $item.File.Extension }
            else { $NewExtension.Text.Trim() }
            if ($ext -and -not $ext.StartsWith('.')) { $ext = '.' + $ext }

            # Build destination path
            $relativeDir = if ($PreserveFolderStructure.IsChecked) { Split-Path $row.SourceRelative -Parent } else { '' }
            $destDir = if ($relativeDir) { Join-Path $dst $relativeDir } else { $dst }
            $row.DestinationPath = Join-Path $destDir ($name + $ext)
            $row.DestinationRelative = $row.DestinationPath.Substring($dst.TrimEnd('\').Length).TrimStart('\')

            # Details: show all field values
            $row.Details = ($values.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join '; '

        }
        catch {
            $row.StatusCode = 'Error'
            $row.Status = (T 'Title_Error')
            $row.Details = $_.Exception.Message
        }

        $script:PreviewRows.Add([pscustomobject]$row)
    }

    # Check for duplicate destinations
    $dupes = @(
        $script:PreviewRows |
        Where-Object { $_.DestinationPath } |
        Group-Object DestinationPath |
        Where-Object { $_.Count -gt 1 } |
        Select-Object -ExpandProperty Name
    )
    foreach ($r in $script:PreviewRows) {
        if ($r.DestinationPath -in $dupes) {
            $r.StatusCode = 'Error'
            $r.Status = (T 'Title_Error')
            $r.Details = (T 'Err_DupDest')
        }
        elseif ($r.DestinationPath -and (Test-Path $r.DestinationPath)) {
            $r.StatusCode = 'Error'
            $r.Status = (T 'Title_Error')
            $r.Details = (T 'Err_FileExists')
        }
    }

    RefreshPreviewGrid
}

function RefreshPreviewGrid {
    $filter = [string]$PreviewFilter.SelectedItem.Content
    $data = switch ($filter) {
        (T 'Cbo_Errors') { @($script:PreviewRows | Where-Object { $_.StatusCode -eq 'Error' }) }
        (T 'Cbo_Ready') { @($script:PreviewRows | Where-Object { $_.StatusCode -eq 'Ready' }) }
        default { @($script:PreviewRows) }
    }
    $PreviewGrid.ItemsSource = $null
    $PreviewGrid.ItemsSource = $data

    $errCount = @($script:PreviewRows | Where-Object { $_.StatusCode -eq 'Error' }).Count
    $PreviewInfo.Text = "$(T 'Msg_Files'): $($script:PreviewRows.Count); $(T 'Msg_Errors'): $errCount"
}

function ExecuteCopy {
    # Rebuild preview to get fresh state
    FullBuildPreview

    #$errors = @($script:PreviewRows | Where-Object { $_.StatusCode -eq 'Error' })
    #if ($errors) {
    #    throw "$(T 'Err_Blocked') $($errors.Count) $(T 'Err_FixPrev')"
    #}

    $mode = if ($ExecutionMode -and $ExecutionMode.SelectedIndex -eq 1) { 'Move' } else { 'Copy' }
    $modeLabel = if ($mode -eq 'Move') { (T 'Action_Move') } else { (T 'Action_Copy') }
    $confirmText = if ($mode -eq 'Move') {
        "$modeLabel $((($script:PreviewRows | Where-Object { $_.StatusCode -eq 'Ready' }).Count)) $(T 'Msg_MoveOrig')"
    }
    else {
        "$modeLabel $((($script:PreviewRows | Where-Object { $_.StatusCode -eq 'Ready' }).Count)) $(T 'Msg_CopyOrig')"
    }

    $total = ($script:PreviewRows | Where-Object { $_.StatusCode -eq 'Ready' }).Count
    $confirm = [Windows.MessageBox]::Show(
        $confirmText,
        (T 'Title_Confirm'), 'YesNo', 'Question'
    )
    if ($confirm -ne 'Yes') { return }

    # Show progress bar
    $ProgressBar.Visibility = 'Visible'
    $ProgressBar.Maximum = $total
    $ProgressBar.Value = 0

    $ok = 0; $fail = 0

    foreach ($r in $script:PreviewRows | Where-Object { $_.StatusCode -eq 'Ready' }) {
        try {
            $dir = Split-Path $r.DestinationPath -Parent
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
            if ($mode -eq 'Move') {
                Move-Item -LiteralPath $r.SourcePath -Destination $r.DestinationPath -ErrorAction Stop -Force
                Log "$(T 'Log_Moved') $($r.SourcePath) -> $($r.DestinationPath)"
            }
            else {
                Copy-Item -LiteralPath $r.SourcePath -Destination $r.DestinationPath -ErrorAction Stop -Force
                Log "$(T 'Log_Copied') $($r.SourcePath) -> $($r.DestinationPath)"
            }
            $ok++
        }
        catch {
            $fail++
            Log "$(T 'Log_CopyErr') $($_.Exception.Message)" 'ERROR'
        }
        $ProgressBar.Value = $ok + $fail
        UpdateUI
    }

    $ProgressBar.Visibility = 'Collapsed'
    if ($mode -eq 'Move') {
        SetStatus "$(T 'Status_Moved') $ok; $(T 'Msg_Errors'): $fail"
    }
    else {
        SetStatus "$(T 'Status_Copied') $ok; $(T 'Msg_Errors'): $fail"
    }
}

#endregion

#region Profile Functions

function RefreshProfiles {
    $items = @(
        Get-ChildItem $script:ProfileRoot -Filter '*.json' -ErrorAction SilentlyContinue |
        ForEach-Object { [pscustomobject]@{ Name = $_.BaseName; Path = $_.FullName } }
    )
    $ProfileList.ItemsSource = $items
}

function SaveProfile {
    $name = [Microsoft.VisualBasic.Interaction]::InputBox(
        (T 'Lbl_ProfileName'), (T 'Title_SaveProfile'), $script:CurrentProfileName
    )
    if (-not $name) { return }

    # Serialize fields with transforms converted to plain arrays
    $fieldsData = @($script:Fields | ForEach-Object {
            $clone = [ordered]@{}
            $_.PSObject.Properties | ForEach-Object { $clone[$_.Name] = $_.Value }
            $clone.Transforms = @($_.Transforms)
            [pscustomobject]$clone
        })

    $obj = [ordered]@{
        SchemaVersion = 2
        Name          = $name
        TokenRegex    = $TokenRegex.Text
        Fields        = $fieldsData
        Mappings      = @($script:Mappings)
        OutputParts   = @($script:OutputParts)
        KeepExtension = [bool]$KeepExtension.IsChecked
        NewExtension  = $NewExtension.Text
    }

    $path = Join-Path $script:ProfileRoot ($name + '.json')
    $obj | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $path -Encoding UTF8

    $script:CurrentProfileName = $name
    $CurrentProfile.Text = $name
    RefreshProfiles
    SetStatus "$(T 'Status_ProfSaved') $name"
}

function LoadProfile([string]$path) {
    $p = Get-Content -LiteralPath $path -Raw -Encoding UTF8 | ConvertFrom-Json

    # Restore fields
    $script:Fields.Clear()
    @($p.Fields) | ForEach-Object {
        # Ensure Transforms is ArrayList
        $transforms = [System.Collections.ArrayList]::new()
        if ($_.PSObject.Properties['Transforms'] -and $_.Transforms) {
            @($_.Transforms) | ForEach-Object { [void]$transforms.Add($_) }
        }
        $_.Transforms = $transforms

        # Backward compatibility: ensure new properties exist
        if (-not $_.PSObject.Properties['IsVirtual']) {
            $_ | Add-Member -NotePropertyName 'IsVirtual' -NotePropertyValue $false -Force
        }
        if (-not $_.PSObject.Properties['Source']) {
            $src = if ($_.IsVirtual) { (T 'Src_Mapping') } else { (T 'Src_Name') }
            $_ | Add-Member -NotePropertyName 'Source' -NotePropertyValue $src -Force
        }
        if (-not $_.PSObject.Properties['PartIndex']) {
            $pi = if ($_.PSObject.Properties['Index']) { $_.Index } else { -1 }
            $_ | Add-Member -NotePropertyName 'PartIndex' -NotePropertyValue $pi -Force
        }
        if (-not $_.PSObject.Properties['DisplayIndex']) {
            $di = if ($_.IsVirtual) { 'V' } else { "$($_.PartIndex)" }
            $_ | Add-Member -NotePropertyName 'DisplayIndex' -NotePropertyValue $di -Force
        }
        if (-not $_.PSObject.Properties['DetectedTypeId']) {
            $_ | Add-Member -NotePropertyName 'DetectedTypeId' -NotePropertyValue (ConvertLegacyTypeId ([string]$_.DetectedType)) -Force
        }
        if (-not $_.PSObject.Properties['CandidateTypes']) {
            $_ | Add-Member -NotePropertyName 'CandidateTypes' -NotePropertyValue @() -Force
        }
        if (-not $_.PSObject.Properties['IsAmbiguous']) {
            $_ | Add-Member -NotePropertyName 'IsAmbiguous' -NotePropertyValue ($_.DetectedTypeId -eq 'Ambiguous') -Force
        }
        if (-not $_.PSObject.Properties['SelectedTypeId']) {
            $_ | Add-Member -NotePropertyName 'SelectedTypeId' -NotePropertyValue 'Auto' -Force
        }
        if (-not $_.PSObject.Properties['SelectedFormat']) {
            $_ | Add-Member -NotePropertyName 'SelectedFormat' -NotePropertyValue $null -Force
        }
        $_.DetectedType = TokenTypeLabel ([string]$_.DetectedTypeId)
        if (-not $_.PSObject.Properties['EffectiveType']) {
            $_ | Add-Member -NotePropertyName 'EffectiveType' -NotePropertyValue (GetEffectiveTypeLabel $_) -Force
        }
        else {
            $_.EffectiveType = GetEffectiveTypeLabel $_
        }

        $script:Fields.Add($_)
    }

    # Restore mappings
    $script:Mappings.Clear()
    @($p.Mappings) | ForEach-Object {
        if (-not $_.PSObject.Properties['Display']) {
            $_ | Add-Member -NotePropertyName 'Display' -NotePropertyValue "$($_.Name): $($_.InputField) → $($_.OutputField)" -Force
        }
        $script:Mappings.Add($_)
    }

    # Restore output parts
    $script:OutputParts.Clear()
    @($p.OutputParts) | ForEach-Object { $script:OutputParts.Add($_) }

    # Restore extension settings
    $KeepExtension.IsChecked = $p.KeepExtension
    $NewExtension.Text = if ($p.NewExtension) { $p.NewExtension } else { '' }
    if ($p.PSObject.Properties['TokenRegex'] -and -not [string]::IsNullOrWhiteSpace([string]$p.TokenRegex)) {
        $TokenRegex.Text = [string]$p.TokenRegex
    }

    # Refresh UI bindings
    $FieldGrid.ItemsSource = $script:Fields
    $MappingList.ItemsSource = $script:Mappings
    $OutputList.ItemsSource = $script:OutputParts
    RefreshFieldSelector
    UpdateOutputExample

    $script:CurrentProfileName = $p.Name
    $CurrentProfile.Text = $p.Name
    SetStatus "$(T 'Status_ProfLoaded') $($p.Name)"
}

#endregion

#region Event Handlers

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
            $result = ShowTransformDialog
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
                $script:Config.Language = $tag
                $script:Config.Version = 2
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