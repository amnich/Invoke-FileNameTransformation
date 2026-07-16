#requires -Version 5.1
<#
.SYNOPSIS
    GUI utility for advanced file renaming and structural transformation.

.DESCRIPTION
    FileNameTransformer.GUI is a WPF-based PowerShell script that enables bulk renaming and copying of files. 
    It parses source filenames using structural patterns, applies value mappings (via CSV dictionaries), 
    and performs text transformations (such as Regex replaces, casing adjustments, and substrings) to generate 
    target filenames based on user-defined templates. The tool offers a live preview grid to catch potential 
    collisions or missing data before executing any file operations. 

    Features:
    - Multi-language UI (Polish, English, German).
    - Saveable and reusable profiles in JSON format.
    - Pattern-based source file parsing.
    - External CSV-based mapping support.
    
.EXAMPLE
    .\Invoke-FileNameTransformation.ps1
    Launches the application GUI.

.NOTES
    Requires PowerShell 5.1 and must be executed in STA (Single-Threaded Apartment) mode.
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

if (Test-Path $script:ConfigPath) {
    try {
        Write-Verbose $script:ConfigPath
        $config = Get-Content -LiteralPath $script:ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
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
        'Chk_Recursive'          = 'Skanuj także podfoldery (zachowaj strukturę)'
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
        'Type_Num'               = 'Liczba'
        'Type_Text'              = 'Tekst'
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
        'Tip_TokenRegexLabel'    = 'Domyślnie: (?&lt;value&gt;[^_\-\s]+)|(?&lt;sep&gt;[_\-\s]+)'
        'Tip_TokenRegex'         = 'Wyrażenie regularne używane do rozbijania nazw plików na bloki (tokeny). Grupy: value = segment wartości, sep = separator.'
    }
    'EN' = @{
        'WinTitle'               = 'File Name Transformer'
        'Folders_Header'         = 'Folders'
        'Source_Folder'          = 'Source folder:'
        'Btn_Browse'             = 'Browse...'
        'Dest_Folder'            = 'Destination folder:'
        'Chk_Recursive'          = 'Scan subfolders (keep structure)'
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
        'Type_Num'               = 'Number'
        'Type_Text'              = 'Text'
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
        'Tip_TokenRegexLabel'    = 'Default: (?&lt;value&gt;[^_\-\s]+)|(?&lt;sep&gt;[_\-\s]+)'
        'Tip_TokenRegex'         = 'Regular expression used to split filenames into blocks (tokens). Groups: value = value segment, sep = separator.'
    }
    'DE' = @{
        'WinTitle'               = 'File Name Transformer'
        'Folders_Header'         = 'Ordner'
        'Source_Folder'          = 'Quellordner:'
        'Btn_Browse'             = 'Durchsuchen...'
        'Dest_Folder'            = 'Zielordner:'
        'Chk_Recursive'          = 'Unterordner scannen (Struktur beibehalten)'
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
        'Type_Num'               = 'Zahl'
        'Type_Text'              = 'Text'
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
        'Tip_TokenRegexLabel'    = 'Standard: (?&lt;value&gt;[^_\-\s]+)|(?&lt;sep&gt;[_\-\s]+)'
        'Tip_TokenRegex'         = 'Regulärer Ausdruck zum Aufteilen von Dateinamen in Blöcke (Token). Gruppen: value = Wertsegment, sep = Trennzeichen.'
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
        <ComboBox x:Name="LanguageSelector" DockPanel.Dock="Right" Width="100" Margin="0,0,15,0" MinHeight="22" Padding="4,2">
          <ComboBoxItem Content="Polski" Tag="PL"/>
          <ComboBoxItem Content="English" Tag="EN"/>
          <ComboBoxItem Content="Deutsch" Tag="DE"/>
        </ComboBox>
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
                    <DataGridTextColumn Header="{t:Col_Type}"      Binding="{Binding DetectedType}" Width="*"/>
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
    $result = @()
    $pattern = if ($TokenRegex -and $TokenRegex.Text) { $TokenRegex.Text } else { '(?<value>[^_\-\s]+)|(?<sep>[_\-\s]+)' }
    $matches_ = [regex]::Matches($name, $pattern)
    foreach ($m in $matches_) {
        $result += [pscustomobject]@{
            Value       = $m.Value
            IsSeparator = $m.Groups['sep'].Success
        }
    }
    $result
}

function TokenType([string]$v) {
    if ($v -match '^\d{8}$') { return (T 'Type_Date_1') }
    if ($v -match '^\d{4}[-_.]\d{2}[-_.]\d{2}$') { return (T 'Type_Date_2') }
    if ($v -match '^\d{2}[-_.]\d{2}[-_.]\d{4}$') { return (T 'Type_Date_3') }
    if ($v -match '^\d+$') { return (T 'Type_Num') }
    return (T 'Type_Text')
}

function ParseNameByTemplate([string]$name, $templateParts) {
    $result = @{}
    $currentString = $name

    for ($i = 0; $i -lt $templateParts.Count; $i++) {
        $p = $templateParts[$i]
        
        if ($p.IsSeparator) {
            $idx = $currentString.IndexOf($p.Value)
            if ($idx -lt 0) {
                throw "Brak separatora '$($p.Value)' w nazwie pliku."
            }
            $value = $currentString.Substring(0, $idx)
            if ($i -gt 0) {
                $result[$i - 1] = $value
            }
            $currentString = $currentString.Substring($idx + $p.Value.Length)
        }
    }
    
    $lastPartIndex = $templateParts.Count - 1
    if (-not $templateParts[$lastPartIndex].IsSeparator) {
        $result[$lastPartIndex] = $currentString
    }
    
    return $result
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
                else { 'T:' + (TokenType $_.Value) + '|' }
            }) -join ''
        $raw += [pscustomobject]@{ Signature = $sig; Parts = $parts; File = $f }
    }

    $script:Patterns = @()
    foreach ($g in $raw | Group-Object Signature) {
        $sample = $g.Group[0]
        $all = $g.Group
        $labels = @()
        for ($i = 0; $i -lt $sample.Parts.Count; $i++) {
            if ($sample.Parts[$i].IsSeparator) {
                $labels += $sample.Parts[$i].Value
                continue
            }
            $uniqueValues = @($all | ForEach-Object { $_.Parts[$i].Value } | Select-Object -Unique)
            $kind = TokenType $sample.Parts[$i].Value
            $labels += if ($uniqueValues.Count -eq 1) { "[$($uniqueValues[0])]" } else { "<$kind>" }
        }
        $script:Patterns += [pscustomobject]@{
            Extension = $ext
            Display   = ($labels -join '')
            Count     = $g.Count
            Items     = $all
            Signature = $g.Name
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
        $type = TokenType $parts[$i].Value

        # Auto-detect role
        $role = if ($type -like 'Data*') { (T 'Role_Date') }
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
    $browse.Text = 'Wybierz...'
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
            $relativeDir = if ($Recursive.IsChecked) { Split-Path $row.SourceRelative -Parent } else { '' }
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
        Name          = $name
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
                $config = @{ Language = $tag }
                $config | ConvertTo-Json | Set-Content -LiteralPath $script:ConfigPath -Encoding UTF8
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
# SIG # Begin signature block
# MIIzcgYJKoZIhvcNAQcCoIIzYzCCM18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDCsdiaDDXmmKGr
# 5m2mVADEV4brEKV65k1gT6YERkwcrKCCLJEwggaCMIIEaqADAgECAhA2wrC9fBs6
# 56Oz3TbLyXVoMA0GCSqGSIb3DQEBDAUAMIGIMQswCQYDVQQGEwJVUzETMBEGA1UE
# CBMKTmV3IEplcnNleTEUMBIGA1UEBxMLSmVyc2V5IENpdHkxHjAcBgNVBAoTFVRo
# ZSBVU0VSVFJVU1QgTmV0d29yazEuMCwGA1UEAxMlVVNFUlRydXN0IFJTQSBDZXJ0
# aWZpY2F0aW9uIEF1dGhvcml0eTAeFw0yMTAzMjIwMDAwMDBaFw0zODAxMTgyMzU5
# NTlaMFcxCzAJBgNVBAYTAkdCMRgwFgYDVQQKEw9TZWN0aWdvIExpbWl0ZWQxLjAs
# BgNVBAMTJVNlY3RpZ28gUHVibGljIFRpbWUgU3RhbXBpbmcgUm9vdCBSNDYwggIi
# MA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQCIndi5RWedHd3ouSaBmlRUwHxJ
# BZvMWhUP2ZQQRLRBQIF3FJmp1OR2LMgIU14g0JIlL6VXWKmdbmKGRDILRxEtZdQn
# Oh2qmcxGzjqemIk8et8sE6J+N+Gl1cnZocew8eCAawKLu4TRrCoqCAT8uRjDeypo
# GJrruH/drCio28aqIVEn45NZiZQI7YYBex48eL78lQ0BrHeSmqy1uXe9xN04aG0p
# KG9ki+PC6VEfzutu6Q3IcZZfm00r9YAEp/4aeiLhyaKxLuhKKaAdQjRaf/h6U13j
# QEV1JnUTCm511n5avv4N+jSVwd+Wb8UMOs4netapq5Q/yGyiQOgjsP/JRUj0MAT9
# YrcmXcLgsrAimfWY3MzKm1HCxcquinTqbs1Q0d2VMMQyi9cAgMYC9jKc+3mW62/y
# Vl4jnDcw6ULJsBkOkrcPLUwqj7poS0T2+2JMzPP+jZ1h90/QpZnBkhdtixMiWDVg
# h60KmLmzXiqJc6lGwqoUqpq/1HVHm+Pc2B6+wCy/GwCcjw5rmzajLbmqGygEgaj/
# OLoanEWP6Y52Hflef3XLvYnhEY4kSirMQhtberRvaI+5YsD3XVxHGBjlIli5u+Nr
# LedIxsE88WzKXqZjj9Zi5ybJL2WjeXuOTbswB7XjkZbErg7ebeAQUQiS/uRGZ58N
# Hs57ZPUfECcgJC+v2wIDAQABo4IBFjCCARIwHwYDVR0jBBgwFoAUU3m/WqorSs9U
# gOHYm8Cd8rIDZsswHQYDVR0OBBYEFPZ3at0//QET/xahbIICL9AKPRQlMA4GA1Ud
# DwEB/wQEAwIBhjAPBgNVHRMBAf8EBTADAQH/MBMGA1UdJQQMMAoGCCsGAQUFBwMI
# MBEGA1UdIAQKMAgwBgYEVR0gADBQBgNVHR8ESTBHMEWgQ6BBhj9odHRwOi8vY3Js
# LnVzZXJ0cnVzdC5jb20vVVNFUlRydXN0UlNBQ2VydGlmaWNhdGlvbkF1dGhvcml0
# eS5jcmwwNQYIKwYBBQUHAQEEKTAnMCUGCCsGAQUFBzABhhlodHRwOi8vb2NzcC51
# c2VydHJ1c3QuY29tMA0GCSqGSIb3DQEBDAUAA4ICAQAOvmVB7WhEuOWhxdQRh+S3
# OyWM637ayBeR7djxQ8SihTnLf2sABFoB0DFR6JfWS0snf6WDG2gtCGflwVvcYXZJ
# JlFfym1Doi+4PfDP8s0cqlDmdfyGOwMtGGzJ4iImyaz3IBae91g50QyrVbrUoT0m
# UGQHbRcF57olpfHhQEStz5i6hJvVLFV/ueQ21SM99zG4W2tB1ExGL98idX8ChsTw
# bD/zIExAopoe3l6JrzJtPxj8V9rocAnLP2C8Q5wXVVZcbw4x4ztXLsGzqZIiRh5i
# 111TW7HV1AtsQa6vXy633vCAbAOIaKcLAo/IU7sClyZUk62XD0VUnHD+YvVNvIGe
# zjM6CRpcWed/ODiptK+evDKPU2K6synimYBaNH49v9Ih24+eYXNtI38byt5kIvh+
# 8aW88WThRpv8lUJKaPn37+YHYafob9Rg7LyTrSYpyZoBmwRWSE4W6iPjB7wJjJpH
# 29308ZkpKKdpkiS9WNsf/eeUtvRrtIEiSJHN899L1P4l6zKVsdrUu1FX1T/ubSrs
# xrYJD+3f3aKg6yxdbugot06YwGXXiy5UUGZvOu3lXlxA+fC13dQ5OlL2gIb5lmF6
# Ii8+CQOYDwXM+yd9dbmocQsHjcRPsccUd5E9FiswEqORvz8g3s+jR3SFCgXhN4wz
# 7NgAnOgpCdUo4uDyllU9PzCCBqcwggSPoAMCAQICEQCQrAhyIP3Fp8RrXMcN9z0G
# MA0GCSqGSIb3DQEBDAUAMFcxCzAJBgNVBAYTAkdCMRgwFgYDVQQKEw9TZWN0aWdv
# IExpbWl0ZWQxLjAsBgNVBAMTJVNlY3RpZ28gUHVibGljIFRpbWUgU3RhbXBpbmcg
# Um9vdCBSNDYwHhcNMjYwMzI1MDAwMDAwWhcNNDEwMzI0MjM1OTU5WjBVMQswCQYD
# VQQGEwJHQjEYMBYGA1UEChMPU2VjdGlnbyBMaW1pdGVkMSwwKgYDVQQDEyNTZWN0
# aWdvIFB1YmxpYyBUaW1lIFN0YW1waW5nIENBIFI0MTCCAiIwDQYJKoZIhvcNAQEB
# BQADggIPADCCAgoCggIBAK7kSqIBrYIcYvlmLVuaA8zw1RfBhkn4G1CoemzjcYtM
# L6yNUvKmwGH7y6/5MuSC1UYP/+9KYDSqvMQt/1hEKHYxMAD9oZpBkoaDQFEKbOJH
# elsKe+BaO0ZcENTKfePcraVkA7wrGAW2XHA5gQCQv4IKori/3PNOXxnDMOk8yIMg
# VrlMeTxqfWJ4XkjT1xc2s9DD7URHWWJOFobTPoWs6mrDFlaY9FlAHDYTfbzvxQHV
# svRmn3W+5ZmCwyk02I8KgGPT/UX4sTz41GiR+ppwUjQXa1+2tEHZbsdAKUtH3OPE
# VtZvlt7atx4h83IdRR8oYi8wjY3OjFKXFecWpQbzzsPxbUKPwMWiTrzwkrFa8dH/
# 1pDKRJt371W62PfqKPayCr/XbnBOlRn8CALSmHnRtGzuAWtTJpcT3BKw6oy8IIL6
# wSbu938F6ZIbRNIc1dKbIJtr4ULN6R5ZfTdNEhwXctqp3RHDbg4fuOl6LjNoaFwj
# ud92EEDhzxFJzE1jqN4csceZIwxOT1aqfsfh0uFQE/lgTBuBs3i6/WL2W1OceWLy
# 3XEdXRK1f0EWCuea6dNfX2RRdjUfk5EltFnJkN2+bWhnK14OPRKcyjOv5hKZ0iV4
# NRNd1+hjtva1rPyzb5Bs7EvFxqEQhgZbOq7qH3nm0rBwA0dxniBOYCFPdu246JCx
# AgMBAAGjggFuMIIBajAfBgNVHSMEGDAWgBT2d2rdP/0BE/8WoWyCAi/QCj0UJTAd
# BgNVHQ4EFgQUOnSlDGfGQlDC/bX8x7spNIL0erkwDgYDVR0PAQH/BAQDAgGGMBIG
# A1UdEwEB/wQIMAYBAf8CAQAwEwYDVR0lBAwwCgYIKwYBBQUHAwgwIwYDVR0gBBww
# GjAIBgZngQwBBAIwDgYMKwYBBAGyMQECAQMIMEwGA1UdHwRFMEMwQaA/oD2GO2h0
# dHA6Ly9jcmwuc2VjdGlnby5jb20vU2VjdGlnb1B1YmxpY1RpbWVTdGFtcGluZ1Jv
# b3RSNDYuY3JsMHwGCCsGAQUFBwEBBHAwbjBHBggrBgEFBQcwAoY7aHR0cDovL2Ny
# dC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVibGljVGltZVN0YW1waW5nUm9vdFI0Ni5w
# N2MwIwYIKwYBBQUHMAGGF2h0dHA6Ly9vY3NwLnNlY3RpZ28uY29tMA0GCSqGSIb3
# DQEBDAUAA4ICAQAy3lJHZvGeA2b43yhzoarvobHVzbfl+RfuPDwej0wCQkYAN6sc
# Tt2GwFe22qbOCv/tllqFlLKQZE+E9jVyuPTbyQHwrM7R0oLapAEDC1+CowsqSRf/
# ptira5Pfd4PoHICnb9coPQtyZmHSQp5y9IGvqWf1qNfq7V2fHZ8DvEQrLUzeoGF9
# BJRYu2OzacW3QQtUum3NOVf0gPRwv6I4991uhncJ6VP4lcpUpHZKB7R3hiIUC09m
# R9KjzPVnXHvL9n2bAwiUECfK5Zezhiw27F2tgi39DETfU8M4n0N6xLgFzsf05M5G
# URX8C9+IX9V6kpmmKtrUzMti4LD66gtmf+mSm934K81NL6YQeMEk1rpYrWPypcW7
# 6Mir6wb1AgseLIHqn/GkeuQm7zOTDf3f5WoX14qVNjZWNHF3JxkutV6ZnhinfCLf
# dv5bnwKWUfceqOajCVntI6uCbHxjBg6SCsexc5AfIGno7gVFvwifT4XONPsSUaJ7
# 1XsJ+EvciVUVnjOO4qxm0fWJTd8a7jP8mc4ZPqwJvQFtOp7+6G+kUJAF0fnE8YgD
# 8uttBReNTa1YmAeFMiqc38e8fI4eLm0zjM/eeGCHasnoqqrbGwcF41iz9HXzFDwN
# 4iD5z3QShp6HRiU3UpTwDJiiXcr0z6pjl7PyzJ3/tmWtGehV7CAfc/WlyzCCBuIw
# ggTKoAMCAQICEQDnTvJVsFBP+tum3/f8i6MVMA0GCSqGSIb3DQEBDAUAMFUxCzAJ
# BgNVBAYTAkdCMRgwFgYDVQQKEw9TZWN0aWdvIExpbWl0ZWQxLDAqBgNVBAMTI1Nl
# Y3RpZ28gUHVibGljIFRpbWUgU3RhbXBpbmcgQ0EgUjQxMB4XDTI2MDMyNTAwMDAw
# MFoXDTM3MDYyNDIzNTk1OVowcjELMAkGA1UEBhMCR0IxFzAVBgNVBAgTDkdyZWF0
# ZXIgTG9uZG9uMRgwFgYDVQQKEw9TZWN0aWdvIExpbWl0ZWQxMDAuBgNVBAMTJ1Nl
# Y3RpZ28gUHVibGljIFRpbWUgU3RhbXBpbmcgU2lnbmVyIFIzNzCCAiIwDQYJKoZI
# hvcNAQEBBQADggIPADCCAgoCggIBALL/w21L3FDZRS0FEXfZuPtUrefibnRSqOT/
# NNyJLOJhXjQfUspqHT+gSSVgbjYThUI/cO+wFQHoOakKQNnSMKdkE8gR69ofXlkk
# 5DAVY/ZlevliOUmlvrw2Vuz4SU28rHfb/Vgd17eqpRIvJuO6XE8vPpPzn4c4iors
# zUF6nwuynKEQ/+rqfDmQbFNKsa+5+Z4f4kXwKdUFxUwUDjQWUhiHRwMlUWGF9N91
# aAvL+9a4sxCgqR/ez8W8HJ/XqvSu1vIeb+J6bDFKKgkv3PJkMMpQ0BsdeXR2FejZ
# XFRXY1w9dZe6gqyMv7px+TpWbYMefECUV0WxoEMgXUk6RKcLo94uUHOdmfZu4Xe8
# ghglyro3/N4VEKTj8dcPPvOBGxFEx1QH6uHKTkWhloGPDScurcZnd8KUtTHl6zml
# QDHM04MwGfsmQViKnYEAYE8RHl5XRE6GTq0ZMb59SIyJX6+CODVic/kW+dhbIS1Z
# 5AP8HaGne/PRG+12QzSneKDJp3Ot+k4GrmmlWT9iy6FNCQ/32K+d4cAZ+Ll7uWbE
# n6Z6gE+tEu7MyZvzWvPNsRKMkcyyflFW1zpRyzutwypALXc9Qg7sFsYERNXa58KZ
# XqU9Onc/tck6+adQJFM9tW8xOnE//P5I4eDj84IGGKqzgUD37ihC+WST3DfY0YBK
# WL0ZaubnAgMBAAGjggGOMIIBijAfBgNVHSMEGDAWgBQ6dKUMZ8ZCUML9tfzHuyk0
# gvR6uTAdBgNVHQ4EFgQUYRDpehKvUcSF1PLPpHQPUM0gr/gwDgYDVR0PAQH/BAQD
# AgbAMAwGA1UdEwEB/wQCMAAwFgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwSgYDVR0g
# BEMwQTAIBgZngQwBBAIwNQYMKwYBBAGyMQECAQMIMCUwIwYIKwYBBQUHAgEWF2h0
# dHBzOi8vc2VjdGlnby5jb20vQ1BTMEoGA1UdHwRDMEEwP6A9oDuGOWh0dHA6Ly9j
# cmwuc2VjdGlnby5jb20vU2VjdGlnb1B1YmxpY1RpbWVTdGFtcGluZ0NBUjQxLmNy
# bDB6BggrBgEFBQcBAQRuMGwwRQYIKwYBBQUHMAKGOWh0dHA6Ly9jcnQuc2VjdGln
# by5jb20vU2VjdGlnb1B1YmxpY1RpbWVTdGFtcGluZ0NBUjQxLmNydDAjBggrBgEF
# BQcwAYYXaHR0cDovL29jc3Auc2VjdGlnby5jb20wDQYJKoZIhvcNAQEMBQADggIB
# AAPqPY3RrM36GXqTpsoHn9TpW5I6z3dkFvc9zPL1W0Egq7j3jtnkbAvRoWeAjGX4
# ZK4sWsmA+u4EJG8okQmybuS/4tDUI5UIQb21n4hG2vihxShrneWB0VoQ2VLQ3jCC
# RmRtAQ+/7H7WVKNiH5Pgl4v2ZTOdPsStzpKnl1YuRrmww/+bcZmLqgk909ywIpZq
# AfubYfbEMYjIckLk90f2mG+L8qaGSS2JJVM02pV5XltZ1fbOFETpRN/PQhwygIv3
# 3qUUjJ1fE4ITgw0McMzRqziWdOJP8ocxxw7qXxz1OdRWCalyL1qvUgAFnZTVdSRi
# MYZKf0wLcQcM/1Xf1W4FW9nff8ERX8RZJGt/TtPuMWmUpf6BCv9Q6o8YyUTtknvZ
# RpSQ0nLttWXdtwsrN2mMgfMuR//gxVrVXvDzCoK/lbiA6dEZOW53lQwBFtEzwE/F
# H8JdhegyYg4PymZOTZrGBEvgsbxe25yEhJ0IdGa1pwCYsarldJhJVMdNcAOU7jyI
# MqHcczav3wtIXp/SwbXZ3xX0mfsLfANSJ47G4qPgx1atb6GIlTaQXzu/p4fTQeAI
# UVzZXT4K984IyfuO7NLjWMtog1wGUpZD98pv+4Mt9Y5bvfPUjaUVjtePy1DVdi0r
# l5ESNYi0zyOmXVxtA5zzxu1H7RdLZOZugT/XjX69rY9bMIIHvDCCBaSgAwIBAgIQ
# G6SpMny8vKlEMbMo1Op93TANBgkqhkiG9w0BAQsFADCBsTELMAkGA1UEBhMCREUx
# EDAOBgNVBAgTB1NhY2hzZW4xEDAOBgNVBAcTB0ZyZWl0YWwxLDAqBgNVBAsTI0Nv
# cnBvcmF0ZSBJVCAtIENlcnRpZmljYXRlIFNlcnZpY2VzMSAwHgYDVQQKExdCR0gg
# RWRlbHN0YWhsd2Vya2UgR21iSDEuMCwGA1UEAxMlQkdIIFN0YW5kYXJkIFJvb3Qg
# Q0EgUlNBIDQwOTYgU0hBLTI1NjAeFw0yNTEyMDUxMzMyMDRaFw0zMjEyMDUxMzQy
# MDNaMIGxMQswCQYDVQQGEwJERTEQMA4GA1UECBMHU2FjaHNlbjEQMA4GA1UEBxMH
# RnJlaXRhbDEsMCoGA1UECxMjQ29ycG9yYXRlIElUIC0gQ2VydGlmaWNhdGUgU2Vy
# dmljZXMxIDAeBgNVBAoTF0JHSCBFZGVsc3RhaGx3ZXJrZSBHbWJIMS4wLAYDVQQD
# EyVCR0ggU3RhbmRhcmQgUm9vdCBDQSBSU0EgNDA5NiBTSEEtMjU2MIICIjANBgkq
# hkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAmWLPLjvA4BAV6LbWDGWRreHQerrTR3K7
# 8FdnB75zoGHYHr9lIcXimJji80OqStuZGkADcAOwy3K0WYlPXeuiQ8yaLWCo2fl1
# WYT97bQ808rS5qSp5RUAuPM5qqFi2FUQ41MekcXn3qeiyy84hOKzFXUM4+dFki7n
# s9J/VfO9CvAxJXkzBDFCO1QLtzbCu+2LV0BJPQBSE5qJq23yB/ijAxsx0anckNpQ
# VP0GrDCNckuv41OlOsLRv6EZFMJKaCYGSrcEQcXLY2jrqLbuqyXdfHNYaGCOr41y
# 5FRjrI14Qtl15345LxY1LweebgOfIpcjDOPruBi1L/k3IeT1g4Zt6DC1Y4PvAiAD
# n11Wxcgl9CXQ0U86/HmHSwAjgUKRKLkVWo6nVLqqGxXcYUMMfoYLTtXsFlvchbxb
# dOhHwUhOAz6L7xoietCNNtAX90BQuYi4hKczK4FN4esgBNsBEMZVDq7jKA8CFdqj
# gQvka6kB+w4ooaOe4AP3Y2JBq0uzXKvE0eRD+07GOKLop/acS7aMrh+khxJxmzzG
# TVZ+sTIxHQvS6hZ6rmgnUCTx5vnzS2g8Cal8mxwXI0zMWca8LYXFDgmuU6HfsV0n
# 31eJZZFaSf3eISxb37eLhxNt3Cdb5RDGREjBNqLBoLEYaQ7buCQavzWwmAOZypMs
# +qj938yJxokCAwEAAaOCAcwwggHIMA4GA1UdDwEB/wQEAwIBBjASBgNVHRMBAf8E
# CDAGAQH/AgECMB0GA1UdDgQWBBQ5TwDIRXn4MWVDK8ocXVyJ8+z4pjBQBgNVHR8E
# STBHMEWgQ6BBhj9odHRwOi8vcGtpLmJnaC1pbnRyYS5uZXQvY3JsL0JHSFN0YW5k
# YXJkUm9vdENBUlNBNDA5NlNIQTI1Ni5jcmwwEAYJKwYBBAGCNxUBBAMCAQAwgcEG
# A1UdIAEB/wSBtjCBszCBsAYKKwYBBAGD+V8CATCBoTBcBggrBgEFBQcCAjBQHk4A
# QgBHAEgAIABTAHQAYQBuAGQAYQByAGQAIABSAG8AbwB0ACAAQwBBACAAQwBlAHIA
# dABpAGYAaQBjAGEAdABlACAAUABvAGwAaQBjAHkwQQYIKwYBBQUHAgEWNWh0dHBz
# Oi8vcGtpLmJnaC1pbnRyYS5uZXQvY3BzL0JHSC1Sb290Q1AtU3RhbmRhcmQucGRm
# MFsGCCsGAQUFBwEBBE8wTTBLBggrBgEFBQcwAoY/aHR0cDovL3BraS5iZ2gtaW50
# cmEubmV0L2FpYS9CR0hTdGFuZGFyZFJvb3RDQVJTQTQwOTZTSEEyNTYuY3J0MA0G
# CSqGSIb3DQEBCwUAA4ICAQBVuSpejkCj3kXH7P4gPh50Ly3NeMgqh3rTjTy4Eb6D
# +sQCAVcbMvGUaS2mFDWSVBqoyrMrw13mqK9EDHuFUCiqMPrODPPFVGYN0+B4cEmG
# Tkf54Cjh3uwQao9nttFzVTb/ezKSBbIpP0ARoad4z8z95cMBd9VEFQPgt4vPwTnd
# uag0pJFQm16RTxqWF17TBFWTu4DBPQyChiZFdSdDLHEHY78VRhH2kRSYmBlY5zwx
# 0cJ4n6mIZwgLu0zES6oYrDvywbWN9zgzffI+NpIVJZtd0WwXXW0lSw9445kaRHCL
# O1hUGfuFDnfiOTpARF3b/tbd9LD5yQsgsn58UbLOZpL0otRX87kNZ9CMiIGmQND4
# XgIZ1A4MG4nGzcKIgiILYrn3+SxM6NAiGttE4rBgVSUJQuWoljwRadStLEIvDiYx
# gIDNb1CWwkb2w4ezDMHe6RbZ9GSAplF2shJb7dFsYEBRAPFUNHL+LsyUxOjW9M00
# zmN9WKIrMODhTzzydg1jDpYcnoJxoHqk0rG9O/wZqvYdlO3f8y0VIkvTXDgk95Dg
# uv2i82NNy5jYAoDTb/gYCJ/SB7+Px8dn9NszC0J4p5266xrkRbTjFgrYPfa70M/k
# KB31wefyjiEXObx+Eyg7YSOky5ndbtg5pMJK758eChqA6DdwDHtssSzeu8LPw7G6
# 9DCCCAYwggXuoAMCAQICE1IAAAAKWeR5eJNiiY4AAAAAAAowDQYJKoZIhvcNAQEL
# BQAwgbExCzAJBgNVBAYTAkRFMRAwDgYDVQQIEwdTYWNoc2VuMRAwDgYDVQQHEwdG
# cmVpdGFsMSwwKgYDVQQLEyNDb3Jwb3JhdGUgSVQgLSBDZXJ0aWZpY2F0ZSBTZXJ2
# aWNlczEgMB4GA1UEChMXQkdIIEVkZWxzdGFobHdlcmtlIEdtYkgxLjAsBgNVBAMT
# JUJHSCBTdGFuZGFyZCBSb290IENBIFJTQSA0MDk2IFNIQS0yNTYwHhcNMjUxMjEy
# MTM1OTA3WhcNMzAxMjEyMTQwOTA3WjCBxTELMAkGA1UEBhMCREUxEDAOBgNVBAgT
# B1NhY2hzZW4xEDAOBgNVBAcTB0ZyZWl0YWwxIDAeBgNVBAoTF0JHSCBFZGVsc3Rh
# aGx3ZXJrZSBHbWJIMSwwKgYDVQQLEyNDb3Jwb3JhdGUgSVQgLSBDZXJ0aWZpY2F0
# ZSBTZXJ2aWNlczFCMEAGA1UEAxM5QkdIIENvZGUgYW5kIERvY3VtZW50IFNpZ25p
# bmcgSXNzdWluZyBDQSBSU0EgMzA3MiBTSEEtMjU2MIIBojANBgkqhkiG9w0BAQEF
# AAOCAY8AMIIBigKCAYEAyi3TjjSUYgUP8zghYc+sqilmYPerbRnUgi4V5m7+YpuA
# YEoRcuWROZcInxY+ak51tJ3xTR3QlwrW48bJu7Wo0iYi2FE+wVb2YFvKN2LvWJxQ
# 4aq1ilH/OiWgD8hjN3a3ezJSksN3dqwMg1M55R+y7/4G72vlOwGW72LZc3WanQAM
# cw+9jhK905tW3koX24JaOQUhLv1ZzyyueI4VUBXuBLS0p5RunenswpXN7hbYEs95
# quCasgE6Uk9Xca1qlMCfAXpZ9LciWUReFfyk7oM7FaEpz0ed9EetXSMD8luzj4lB
# La+aYClh9XHwbwuBkiXigx4TyAKZDuwDH65mBTetNmUwptSihTIpsUOmYB+LFPXv
# 4mZ11bt2onqU1/ZnHsmUmDBpbHHJxqhKGRe0K90Vu1hPCoN56dUEQ9IiQ/eGW2zp
# zVvi59ET1GDZPXvUYGLuMFiZ6/xz51Xym+IAMFbsX9c//a5k1L60v8knYhiR+Wce
# ba3AUSiZNIXFVQ9actFRAgMBAAGjggJ/MIICezAQBgkrBgEEAYI3FQEEAwIBADAd
# BgNVHQ4EFgQUSRMN98P8m1nKd+HPrniby6Uw8w4wggEDBgNVHSABAf8EgfgwgfUw
# gfIGCisGAQQBg/lfAgQwgeMwgYQGCCsGAQUFBwICMHgedgBCAEcASAAgAEMAbwBk
# AGUAIABhAG4AZAAgAEQAbwBjAHUAbQBlAG4AdAAgAFMAaQBnAG4AaQBuAGcAIABJ
# AHMAcwB1AGkAbgBnACAAQwBBACAAQwBlAHIAdABpAGYAaQBjAGEAdABlACAAUABv
# AGwAaQBjAHkwWgYIKwYBBQUHAgEWTmh0dHBzOi8vcGtpLmJnaC1pbnRyYS5uZXQv
# Y3BzL0JHSC1TdGFuZGFyZElzc3VpbmdDUC1Db2RlYW5kRG9jdW1lbnRTaWduaW5n
# LnBkZjA1BgNVHSUELjAsBggrBgEFBQcDAwYIKwYBBQUHAwkGCisGAQQBgjcUAgEG
# CisGAQQBgjcKAwwwGQYJKwYBBAGCNxQCBAweCgBTAHUAYgBDAEEwCwYDVR0PBAQD
# AgGGMBIGA1UdEwEB/wQIMAYBAf8CAQAwHwYDVR0jBBgwFoAUOU8AyEV5+DFlQyvK
# HF1cifPs+KYwUAYDVR0fBEkwRzBFoEOgQYY/aHR0cDovL3BraS5iZ2gtaW50cmEu
# bmV0L2NybC9CR0hTdGFuZGFyZFJvb3RDQVJTQTQwOTZTSEEyNTYuY3JsMFsGCCsG
# AQUFBwEBBE8wTTBLBggrBgEFBQcwAoY/aHR0cDovL3BraS5iZ2gtaW50cmEubmV0
# L2FpYS9CR0hTdGFuZGFyZFJvb3RDQVJTQTQwOTZTSEEyNTYuY3J0MA0GCSqGSIb3
# DQEBCwUAA4ICAQCUVpwjlOewn05pi1HhO/ceIv0CZpQcKzsEIL2HMS8fmjwA2vb3
# x04dPWK3f1oFLCbXl0yuclSqZeFBGq8/RT3tLB9pwT3ffEgHjp2YvkLr7TIfKcLe
# 7H4bLegfGBegrKQmDHC7CNbvuF4Mk29nakqccNYHOrK6fo/xkOqFjDDLlCLaVVn2
# DhWBcuQAN77ITFe+hkFuGIcR25TwCCpSL6f6LDaa0ZfRFevwXdyz9aeeMKaZS5Hj
# q4iNgKJH7U3S8n3TBNpEzMnqJGMUfVg27IrVF1swOzNGtERHm1cUpYIk8+45CwfI
# quej05waIqwRgAJO+5xbTqqAS6NbmzGm2ijEUzOarInAxZ216fCMTMOZsHhMdZ3M
# 2mXplCRN9/RuENXuf9OY1XY5sCByMaZXxyTzNmM4PV9d++C9vleZXtHQm4IFD0Co
# E/Rpx2c5zibY6emctkr2FwTkficznHEKF0U07iAgLdTZHxu7KosDre+mSqIoXOwQ
# ToSQCkfvw3V9HTltdx83NcUQFQAAQHuabvUQnPNUT8KmHZm47PhZy8Jo/ZIE+bbZ
# N+p/IkINCAJwONLuuhPkW9/rNWHvAd5jchScAF1hT2T+ncvBcaywM01lstxW6Duo
# Y0OfqVTubTmvu1ezw88vp5qpJbgWfrdeJk5ZIhF3ZnznbNxrj5n+Unn0rjCCCKww
# ggcUoAMCAQICExMAAAAQ/9RfhbDcNocAAAAAABAwDQYJKoZIhvcNAQELBQAwgcUx
# CzAJBgNVBAYTAkRFMRAwDgYDVQQIEwdTYWNoc2VuMRAwDgYDVQQHEwdGcmVpdGFs
# MSAwHgYDVQQKExdCR0ggRWRlbHN0YWhsd2Vya2UgR21iSDEsMCoGA1UECxMjQ29y
# cG9yYXRlIElUIC0gQ2VydGlmaWNhdGUgU2VydmljZXMxQjBABgNVBAMTOUJHSCBD
# b2RlIGFuZCBEb2N1bWVudCBTaWduaW5nIElzc3VpbmcgQ0EgUlNBIDMwNzIgU0hB
# LTI1NjAeFw0yNjAzMjYwNzE5NTZaFw0yODAzMjUwNzE5NTZaMBYxFDASBgNVBAMT
# C01uaWNoLCBBZGFtMIIBojANBgkqhkiG9w0BAQEFAAOCAY8AMIIBigKCAYEAsoYE
# C86WcNB53nIx6rXaZbArOegShll++ECrXwPDj4Qn/vn+Pcedho6Oyt8d174he4Ar
# auVj4wPCZEiLfrY8wZy74U8pc5uFF0UGOWWKg5XqsmqCrRDPtY6JmtAlmPMrioUI
# zMVGxWxsN/psb+KNZNrbvxWusSTf5X120VBztvHH5jZB2tmsEhMFtsZKpchGSj+u
# TqahscLbI0M7g8otca5+3uGZeJ0UB1y22/g5Xq/hem/eggK4NPdu5GQUw1SO+bzj
# xqr5sc/TWxrqdeT5UJ/7EdF/Y+q57Zdpmo4heKwe0+V7fj64rphjFiXqYYLWGWd+
# 9jC//pO1kLNpW8McdHLVFxctO+zOJLBG7MV52pbUv1Y+dSAEpL8H1xWtCjwGCmbn
# Ix9pZjaOnK24tlJ4IsTjXgf6bmykpT61ma8sANMa0X2K6/eJEqp4aJrDPUEvz1Jf
# Bp4VILJGXndImxo5a3hjTO/V1O0O4GUXACNhgustVhp7jNCevllJLu/7oOwlAgMB
# AAGjggRBMIIEPTA9BgkrBgEEAYI3FQcEMDAuBiYrBgEEAYI3FQiCy6hAgrTpX4aB
# nzmBtK8kh7KtbhyC4KYGgvywewIBZAIBETATBgNVHSUEDDAKBggrBgEFBQcDAzAO
# BgNVHQ8BAf8EBAMCB4AwGwYJKwYBBAGCNxUKBA4wDDAKBggrBgEFBQcDAzAdBgNV
# HQ4EFgQU5Q+l/z48mBbWUXsQsCHBwxFnOrMwHwYDVR0jBBgwFoAUSRMN98P8m1nK
# d+HPrniby6Uw8w4wggFlBgNVHR8EggFcMIIBWDCCAVSgggFQoIIBTIaB92xkYXA6
# Ly8vQ049QkdIJTIwQ29kZSUyMGFuZCUyMERvY3VtZW50JTIwU2lnbmluZyUyMElz
# c3VpbmclMjBDQSUyMFJTQSUyMDMwNzIlMjBTLTA0MDY0LENOPXN6ZnRsc2Nhc3Jz
# YTAzLENOPUNEUCxDTj1QdWJsaWMlMjBLZXklMjBTZXJ2aWNlcyxDTj1TZXJ2aWNl
# cyxDTj1Db25maWd1cmF0aW9uLERDPWJnaCxEQz1pbnRyYT9jZXJ0aWZpY2F0ZVJl
# dm9jYXRpb25MaXN0P2Jhc2U/b2JqZWN0Q2xhc3M9Y1JMRGlzdHJpYnV0aW9uUG9p
# bnSGUGh0dHA6Ly9wa2kuYmdoLWludHJhLm5ldC9jcmwvQkdIQ29kZWFuZERvY3Vt
# ZW50U2lnbmluZ0lzc3VpbmdDQVJTQTMwNzJTSEEyNTYuY3JsMIIBiwYIKwYBBQUH
# AQEEggF9MIIBeTCB6AYIKwYBBQUHMAKGgdtsZGFwOi8vL0NOPUJHSCUyMENvZGUl
# MjBhbmQlMjBEb2N1bWVudCUyMFNpZ25pbmclMjBJc3N1aW5nJTIwQ0ElMjBSU0El
# MjAzMDcyJTIwUy0wNDA2NCxDTj1BSUEsQ049UHVibGljJTIwS2V5JTIwU2Vydmlj
# ZXMsQ049U2VydmljZXMsQ049Q29uZmlndXJhdGlvbixEQz1iZ2gsREM9aW50cmE/
# Y0FDZXJ0aWZpY2F0ZT9iYXNlP29iamVjdENsYXNzPWNlcnRpZmljYXRpb25BdXRo
# b3JpdHkwXAYIKwYBBQUHMAKGUGh0dHA6Ly9wa2kuYmdoLWludHJhLm5ldC9haWEv
# QkdIQ29kZWFuZERvY3VtZW50U2lnbmluZ0lzc3VpbmdDQVJTQTMwNzJTSEEyNTYu
# Y3J0MC4GCCsGAQUFBzABhiJodHRwOi8vb2NzcC5wa2kuYmdoLWludHJhLm5ldC9v
# Y3NwMDMGA1UdEQQsMCqgKAYKKwYBBAGCNxQCA6AaDBhhbW5pY2hAYmdoLWVkZWxz
# dGFobC5jb20wTQYJKwYBBAGCNxkCBEAwPqA8BgorBgEEAYI3GQIBoC4ELFMtMS01
# LTIxLTg1NDg5MzY2Mi0yMjUzOTAzNDcwLTk0MTgzOTg4MC0yMzY0MA0GCSqGSIb3
# DQEBCwUAA4IBgQBZrSjlosMUSWsJAOc05Z0Yhwbhp3jhmVRyo6v2iX5RLQm/Y42b
# A5VdhStM2m0S7oPqyV9DdfLKGHCsmuJ2+irnWlW6aUzg1bBu0g29FfA2C8tCMtF5
# FRznFQK/erjfq24ZdUZSjBvy8SzpIUgPX+wba10OrMVA/bWZ8R7zuYuHlEzbbpKo
# NKMdS3bYsygr1XeYXPQh0j44klu5PCjDYHwBrPuBzmziEPVol3xmRtMretDHMSvU
# cvS5iY2YgOWLPhNMduSHGOnost9z3E7FqPAVq/6CP4iJDICE4cxfeCC8Hyjk+DGs
# cwFXUajaAfmmqQxE0TdaFyeR44LSredie5XPZq5oiSpEXJ9P/ZTQDSGfKI0vweXc
# fAc0QikVhJAuKtW1KdhH1YONqkLjW+Uz929QQta2Z3SSnFC1jXZ5TgmuphvbmqdL
# BzsUG9p6wtN2yTHYJK0S3XWKzikH4nwOqfXF+WCZCRRkMUdAUP/4MZZzhhl7x8vj
# YJxyTQwc/lLado0xggY3MIIGMwIBATCB3TCBxTELMAkGA1UEBhMCREUxEDAOBgNV
# BAgTB1NhY2hzZW4xEDAOBgNVBAcTB0ZyZWl0YWwxIDAeBgNVBAoTF0JHSCBFZGVs
# c3RhaGx3ZXJrZSBHbWJIMSwwKgYDVQQLEyNDb3Jwb3JhdGUgSVQgLSBDZXJ0aWZp
# Y2F0ZSBTZXJ2aWNlczFCMEAGA1UEAxM5QkdIIENvZGUgYW5kIERvY3VtZW50IFNp
# Z25pbmcgSXNzdWluZyBDQSBSU0EgMzA3MiBTSEEtMjU2AhMTAAAAEP/UX4Ww3DaH
# AAAAAAAQMA0GCWCGSAFlAwQCAQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKEC
# gAAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwG
# CisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEILwKL/Wr5F0y1tWjxzg0zl99srex
# nfKcFLITUVbGrL2QMA0GCSqGSIb3DQEBAQUABIIBgDlb4mApiTN3NxyKN5L8twTH
# ORY3u41P8r0c9zBTgh2EDQOxuuFL3CVYZYCSODjPdt5l/p5U9glIIADag6SR1qYB
# 7xSB6fR7mF7bz7XpwlVdg+kUwcADEv9RhuCGj2Ul608wCF9rgOIWupHwnkLEdwqn
# P63wg5LJdh/9f5eJHWGqTRz7bMwB339n8GO2P4yiZbErqFQW9o3Zgr0OKo8wCpLv
# p13hEgwpF9VWZBRjqonI4RjVNmpB2GlDuiAI+xzKsvjzzN/OlO2tA8lyM1MODP1O
# ym1rYA+yMp+80+gBBAb54Gil80o8iuMsJEFTvc5KDnLfgftE40IvAhfDyH9x16g+
# UzcZX+e+fNnDYKitHi1MOBs7eWCAlUIngw3t6/aIJhbI+SxWwla8lfYIx3+FXebT
# IvBP0meuYAAKsnkj1XLChELSKv3/JB6dk49PjpNJKP/ixyUpSholaX8aHYrtfzIr
# ta5fQ9XclWM+qPlYaAkKSEB2J8LsKce3exwMMfY8BKGCAyMwggMfBgkqhkiG9w0B
# CQYxggMQMIIDDAIBATBqMFUxCzAJBgNVBAYTAkdCMRgwFgYDVQQKEw9TZWN0aWdv
# IExpbWl0ZWQxLDAqBgNVBAMTI1NlY3RpZ28gUHVibGljIFRpbWUgU3RhbXBpbmcg
# Q0EgUjQxAhEA507yVbBQT/rbpt/3/IujFTANBglghkgBZQMEAgIFAKB5MBgGCSqG
# SIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTI2MDcxNjExMzcy
# M1owPwYJKoZIhvcNAQkEMTIEMGnc2jawWcU31m1CUxgloVA8dZtreN95GDyXoQKb
# FQgwoDcgWoO9tujrzzvCekrvLjANBgkqhkiG9w0BAQEFAASCAgCe8sGMMsGT5IwN
# InRMV17l3oQYZWvkBl4shoaGnLpaff0Z/mDYWs0U37w88/zjPWOgwomFDJ7gHyXh
# 799o6DHQl7NrdXnlYDh0QQK8rsjhF2YQMwyog9H0Lcan1W3UBQTfcyVhD7nvcwkf
# FGufLQIIDrykX69NoapdSWFhbP6ousE05ROdSchRIoRZMEWjDCaFgoUJlUE+vJzc
# mTbcxutu3cyf+/OFzIuDBUfNXK0Xst+xViiiUj9IfezkFwemu6gG5/xiL2dQ9BLB
# 4LdXHGG8hFlqfzNGhNtbHPsPx6QDVbPtZcx4DYMndulo8qUcR3HDfqUDLeUvQLhq
# dV4PLnM5tNLb61miRcFbnrl8+Q27xamiMaJ3W9UmFg14eBDIqo7g1pFUXRG7SPIX
# BV6j4OHFydvWlJuUegyTRR6OB5o7a258tIOXHGDL9og4PC0tt8WFX6B7YoBtYhRE
# HCcltb+VZmvEfY7M+Bh5ztA2yyIDb/rVhgZBttyWvQeVXgNAfRC/FvC4jiSSAtvv
# 5AahSra58YImppgTaU3LXqX9CYKwRHOznZ7XsfLtkPpPMX+FTYKrnyXk4VcnXUtj
# Db4MdX4j4md7goYEo2JE7SezuvK5+QcDWxwqeRL3d/OaY0sP3RyBGeG6IXelcVyr
# pfks6peE8iI1W5GARY63wMxJ7lK8Wg==
# SIG # End signature block
