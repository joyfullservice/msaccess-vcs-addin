VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTranslation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsTranslation
' Author    : Adam Waller
' Date      : 5/15/2021
' Purpose   : Used for the translation of strings to different languages, similar to
'           : the gettext toolset.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit


' Enum for sections
Private Enum eFileSection
    efsUnknown
    efsHeader
    efsContext
    efsReference
    efsMessageId
    efsMessageString
End Enum

' Cache strings to dictionary objects so we don't have to do database lookups
' each time we need to return translated strings
Private Type udtThis
    dStrings As Dictionary      ' Untranslated strings (value of string ID)
    dTranslation As Dictionary  ' Translated strings in local language (key of string ID)
    dTranslated As Dictionary   ' Translated strings (key of translated string)
    CurrentLanguage As String   ' Identifier of current language (i.e. "en_US")
End Type
Private this As udtThis


'---------------------------------------------------------------------------------------
' Procedure : GetLanguageName
' Author    : Adam Waller
' Date      : 5/26/2021
' Purpose   : Return the display name for the language. (Add new languages here)
'---------------------------------------------------------------------------------------
'
Private Function GetLanguageName(strIdentifier As String) As String
    Select Case strIdentifier
        Case "en_US":   GetLanguageName = "English"
        Case "pt_BR":   GetLanguageName = "Brazilian Portuguese"
        Case "en_TEST": GetLanguageName = "English (Localization Testing)"
        ' For undefined languages, use the identifier for now.
        Case Else:      GetLanguageName = strIdentifier
    End Select
End Function


'---------------------------------------------------------------------------------------
' Procedure : T
' Author    : Adam Waller
' Date      : 5/15/2021
' Purpose   : Return the translated version of the string.
'---------------------------------------------------------------------------------------
'
Public Function T(strText As String, Optional strReference As String, _
    Optional strContext As String, Optional strComments As String, _
    Optional var0, Optional var1, Optional var2, Optional var3, Optional var4, _
    Optional var5, Optional var6, Optional var7, Optional var8, Optional var9) As String

    Dim strNew As String
    Dim strKey As String

    ' Skip processing if no value is passed
    If strText = vbNullString Then Exit Function

    ' Make sure the language has been initialized
    CheckInit

    ' Fall back to English (untranslated)
    strNew = strText

    ' Check for the master string
    strKey = BuildKey(strContext, strText)
    If this.dStrings.Exists(strKey) Then
        If this.dTranslation.Exists(this.dStrings(strKey)) Then
            ' Use translated string
            strNew = this.dTranslation(this.dStrings(strKey))
        End If
    Else
        If this.dTranslated.Exists(strText) Then
            ' We have already translated this string. Don't add the translation
            ' as a new string in our translation database.
        Else
            ' Add to master list of strings (no translation exists)
            this.dStrings.Add strKey, SaveString(strText, strReference, strContext, strComments)
            this.dTranslated.Add strText, vbNullString
        End If
    End If

    ' Fill in any replacements
    If InStr(1, strNew, "{") > 0 Then
        If Not IsMissing(var0) Then strNew = Replace(strNew, "{0}", CStr(var0))
        If Not IsMissing(var1) Then strNew = Replace(strNew, "{1}", CStr(var1))
        If Not IsMissing(var2) Then strNew = Replace(strNew, "{2}", CStr(var2))
        If Not IsMissing(var3) Then strNew = Replace(strNew, "{3}", CStr(var3))
        If Not IsMissing(var4) Then strNew = Replace(strNew, "{4}", CStr(var4))
        If Not IsMissing(var5) Then strNew = Replace(strNew, "{5}", CStr(var5))
        If Not IsMissing(var6) Then strNew = Replace(strNew, "{6}", CStr(var6))
        If Not IsMissing(var7) Then strNew = Replace(strNew, "{7}", CStr(var7))
        If Not IsMissing(var8) Then strNew = Replace(strNew, "{8}", CStr(var8))
        If Not IsMissing(var9) Then strNew = Replace(strNew, "{9}", CStr(var9))
    End If

    ' Maintain index of fully translated strings, so we don't attempt to translate
    ' a string twice.
    If Not this.dTranslated.Exists(strNew) Then this.dTranslated.Add strNew, vbNullString

    ' Return translated string
    T = strNew

End Function


'---------------------------------------------------------------------------------------
' Procedure : ApplyTo
' Author    : Adam Waller
' Date      : 5/17/2021
' Purpose   : Apply language translation to a form object (From English values)
'           : This also includes placeholder replacements. (Use tag property for {0})
'---------------------------------------------------------------------------------------
'
Public Sub ApplyTo(frmObject As Form)

    Dim ctl As Control
    Dim ctl2 As Control
    Dim strReference As String
    Dim strName As String
    Dim dReferences As Dictionary

    Set dReferences = New Dictionary

    ' Loop through all controls
    For Each ctl In frmObject.Controls

        ' Reset reference string
        strReference = frmObject.Name & "." & ctl.Name

        ' Only check certain types of controls
        Select Case TypeName(ctl)
            Case "Label"

                ' Check for associated control
                ' (It is easier to go from the object to the label, but not
                '  all labels may have objects, so we loop through other controls
                On Error Resume Next
                For Each ctl2 In frmObject.Controls
                    strName = vbNullString
                    strName = ctl2.Controls(0).Name
                    If strName = ctl.Name Then
                        ' Found associated label
                        ' Add extended context
                        strReference = strReference & "(" & ctl2.Name & ")"
                        Exit For
                    End If
                Next ctl2
                If DebugMode(False) Then On Error GoTo 0 Else On Error Resume Next

                ' Translation caption
                ctl.Caption = T(ctl.Caption, strReference)

            Case "TextBox"
                ' Nothing to translate

            Case "CommandButton"
                ctl.Caption = T(ctl.Caption, strReference)

            Case "Page"
                ' Tab control page caption
                strReference = frmObject.Name & "." & ctl.Parent.Name & "." & ctl.Name
                ctl.Caption = T(ctl.Caption, strReference)

            Case Else
                'Debug.Print TypeName(ctl)

        End Select

        ' Maintain list of references so we can purge unused entries
        If Len(strReference) Then If Not dReferences.Exists(strReference) Then dReferences.Add strReference, vbNullString

    Next ctl

    ' Other properties
    frmObject.Caption = T(frmObject.Caption, frmObject.Name & ".Caption")
    dReferences.Add frmObject.Name & ".Caption", vbNullString

    ' Purge any unused references
    If Me.Contribute Then PurgeStrings frmObject.Name & ".", dReferences
    PurgeStrings frmObject.Name & ".", dReferences

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TranslateStrings
' Author    : Adam Waller
' Date      : 10/4/2024
' Purpose   : Translate a dictionary of key-value pairs where the key is the string to
'           : translate, and the value is the translation. This can be used to translate
'           : strings from an external component, like a menu handler.
'---------------------------------------------------------------------------------------
'
Public Function TranslateStrings(dStrings As Dictionary) As Dictionary
    Dim varKey As Variant
    Set TranslateStrings = New Dictionary
    With TranslateStrings
        For Each varKey In dStrings.Keys
            .Add varKey, T(CStr(varKey))
        Next varKey
    End With
End Function


'---------------------------------------------------------------------------------------
' Procedure : ExportTemplate
' Author    : Adam Waller
' Date      : 11/7/2024
' Purpose   : Exports the template file for the English strings used for translation.
'---------------------------------------------------------------------------------------
'
Public Sub ExportTemplate()
    WriteFileNoBom BuildFileContent("en_US"), FSO.BuildPath(TranslationsPath, PROJECT_NAME & ".pot")
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ExportTranslations
' Author    : Adam Waller
' Date      : 5/19/2021
' Purpose   : Export translations to files
'---------------------------------------------------------------------------------------
'
Public Sub ExportTranslations()

    Dim dbs As Database
    Dim rst As Recordset
    Dim strLanguage As String
    Dim strFolder As String
    Dim strFile As String

    strFolder = TranslationsPath
    If strFolder = vbNullString Then Exit Sub

    Set dbs = CodeDb
    Set rst = dbs.OpenRecordset("tblLanguages", dbOpenSnapshot)
    With rst
        Do While Not .EOF
            strLanguage = Nz(!ID)
            If Len(strLanguage) And strLanguage <> "en_US" Then
                ' Build translation work file
                strFile = FSO.BuildPath(strFolder, strLanguage & ".po")
                WriteFileNoBom BuildFileContent(strLanguage), strFile
            End If
            .MoveNext
        Loop
        .Close
    End With

End Sub


'---------------------------------------------------------------------------------------
' Procedure : LoadTranslations
' Author    : Adam Waller
' Date      : 5/19/2021
' Purpose   : Load translation data from translation files
'---------------------------------------------------------------------------------------
'
Public Sub LoadTranslations()

    Dim strBase As String
    Dim strFile As String
    Dim dFiles As Dictionary
    Dim varFile As Variant

    strBase = AddSlash(TranslationsPath)

    ' Load English translation
    strFile = strBase & PROJECT_NAME & ".pot"
    If FSO.FileExists(strFile) Then ImportTranslation strFile

    ' Load other translations
    Set dFiles = GetFileList(strBase, "*.po")
    For Each varFile In dFiles.Keys
        strFile = strBase & varFile
        ImportTranslation strFile
    Next varFile

    ' Reload current translation
    LoadLanguage Me.Language

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TranslationsPath
' Author    : Adam Waller
' Date      : 3/22/2024
' Purpose   : Get translations path from the following locations:
'           :   1. Saved translation path (If contribute is on)
'           :   2. Local translations path, relative to CodeDB
'---------------------------------------------------------------------------------------
'
Public Property Get TranslationsPath() As String

    Dim strPath As String

    '   1. Saved translations path
    If Me.Contribute Then
        strPath = GetSetting(PROJECT_NAME, "Language", "Translation Path", vbNullString)
        If HasTranslationFiles(strPath) Then
            TranslationsPath = strPath
            Exit Property
        End If
    End If

    '   2. Local translations folder
    strPath = FSO.BuildPath(CodeProject.Path, "Translation")
    If HasTranslationFiles(strPath) Then
        TranslationsPath = strPath
    End If

End Property


'---------------------------------------------------------------------------------------
' Procedure : StringCount
' Author    : Adam Waller
' Date      : 11/7/2024
' Purpose   : Return count of strings from table. (Used to verify that we have
'           : successfully loaded strings.)
'---------------------------------------------------------------------------------------
'
Public Property Get StringCount() As Long
    Dim dbs As Database
    Dim rst As Recordset
    Set dbs = CodeDb
    Set rst = dbs.OpenRecordset("select count(*) from tblStrings", dbOpenSnapshot, dbReadOnly)
    StringCount = Nz(rst.Fields(0))
    rst.Close
End Property


'---------------------------------------------------------------------------------------
' Procedure : ShowStrings
' Author    : Adam Waller
' Date      : 11/8/2024
' Purpose   : Open the strings table directly for review.
'---------------------------------------------------------------------------------------
'
Public Sub ShowStrings()
    DoCmd.OpenTable "tblStrings"
    DoCmd.OpenTable "tblTranslation"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : TranslationsPath
' Author    : Adam Waller
' Date      : 3/23/2024
' Purpose   : Set the translations path
'---------------------------------------------------------------------------------------
'
Public Property Let TranslationsPath(strFolderPath As String)
    SaveSetting PROJECT_NAME, "Language", "Translation Path", strFolderPath
End Property


'---------------------------------------------------------------------------------------
' Procedure : HasTranslationFiles
' Author    : Adam Waller
' Date      : 3/22/2024
' Purpose   : Returns true if the folder contains the master template file.
'---------------------------------------------------------------------------------------
'
Private Function HasTranslationFiles(strFolder As String) As Boolean
    If FSO.FolderExists(strFolder) Then
        ' Check for the template file "MSAccessVCS.pot"
        HasTranslationFiles = FSO.FileExists(FSO.BuildPath(strFolder, PROJECT_NAME & ".pot"))
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : Contribute
' Author    : Adam Waller
' Date      : 5/19/2021
' Purpose   : Whether the user desires to contribute to translations
'---------------------------------------------------------------------------------------
'
Public Property Get Contribute() As Boolean
    Contribute = GetSetting(PROJECT_NAME, "Language", "Contribute To Translations", False)
End Property
Public Property Let Contribute(blnContributeToTranslations As Boolean)
    SaveSetting PROJECT_NAME, "Language", "Contribute To Translations", blnContributeToTranslations
End Property


'---------------------------------------------------------------------------------------
' Procedure : Language
' Author    : Adam Waller
' Date      : 5/19/2021
' Purpose   : Selected language
'---------------------------------------------------------------------------------------
'
Public Property Get Language() As String
    Language = GetSetting(PROJECT_NAME, "Language", "Language", "en_US")
End Property
Public Property Let Language(strLanguage As String)
    SaveSetting PROJECT_NAME, "Language", "Language", strLanguage
End Property


'---------------------------------------------------------------------------------------
' Procedure : PurgeStrings
' Author    : Adam Waller
' Date      : 8/23/2024
' Purpose   : Purges translation strings from form controls that no longer exist
'           : (Note that renamed controls are also purged and added as untranslated)
'---------------------------------------------------------------------------------------
'
Private Sub PurgeStrings(strObjectPrefix As String, dReferences As Dictionary)

    Dim dbs As DAO.Database
    Dim rst As DAO.Recordset
    Dim strSql As String

    Set dbs = CodeDb
    strSql = "select * from tblStrings where Reference like """ & strObjectPrefix & "*"""
    Set rst = dbs.OpenRecordset(strSql, dbOpenDynaset)
    With rst
        Do While Not .EOF
            If Len(Nz(!Reference)) > 0 Then
                If Not dReferences.Exists(Nz(!Reference)) Then
                    ' Delete orphaned string entry
                    ' (Translations will be removed automatically when
                    '  the translation files are rewritten.)
                    Debug.Print "Removing translation string: " & Nz(!Reference) & ": " & Nz(!msgid)
                    .Delete
                End If
            End If
            .MoveNext
        Loop
        .Close
    End With

End Sub


'---------------------------------------------------------------------------------------
' Procedure : BuildKey
' Author    : Adam Waller
' Date      : 8/21/2024
' Purpose   : Build key from context and string
'---------------------------------------------------------------------------------------
'
Private Function BuildKey(strContext As String, strText As String) As String
    If Len(strContext) > 0 Then
        BuildKey = strContext & "|" & strText
    Else
        BuildKey = strText
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : SaveString
' Author    : Adam Waller
' Date      : 5/15/2021
' Purpose   : Save the string to the database table and return ID
'---------------------------------------------------------------------------------------
'
Private Function SaveString(strText As String, strReference As String, _
    strContext As String, strComments As String) As Long

    Dim dbs As Database
    Dim rst As Recordset

    Set dbs = CodeDb
    Set rst = dbs.OpenRecordset("tblStrings")

    With rst
        .AddNew
            !msgid = strText
            !Reference = ZN(strReference)
            !Context = ZN(strContext)
            !Comments = ZN(strComments)
        .Update
        .Bookmark = .LastModified
        SaveString = !ID
        .Close
    End With

End Function


'---------------------------------------------------------------------------------------
' Procedure : SetLanguage
' Author    : Adam Waller
' Date      : 5/15/2021
' Purpose   : Set the current language
'---------------------------------------------------------------------------------------
'
Public Sub SetLanguage(strLanguage As String)
    LoadLanguage strLanguage
End Sub


'---------------------------------------------------------------------------------------
' Procedure : CheckInit
' Author    : Adam Waller
' Date      : 5/15/2021
' Purpose   : Ensure that the language strings have been loaded
'---------------------------------------------------------------------------------------
'
Private Sub CheckInit()
    If this.CurrentLanguage = vbNullString Then
        Set this.dStrings = New Dictionary
        Set this.dTranslation = New Dictionary
        Set this.dTranslated = New Dictionary
        LoadLanguage GetCurrentLanguage
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : LoadLanguage
' Author    : Adam Waller
' Date      : 5/15/2021
' Purpose   : Loads the language entries into the dictionary objects.
'---------------------------------------------------------------------------------------
'
Private Sub LoadLanguage(strLanguage As String)

    Dim dbs As Database
    Dim rst As Recordset
    Dim strKey As String
    Dim strSql As String
    Dim strPath As String

    this.CurrentLanguage = strLanguage
    Set this.dStrings = New Dictionary
    Set this.dTranslation = New Dictionary
    Set this.dTranslated = New Dictionary

    ' Load strings and translations
    Set dbs = CodeDb
    strSql = "select * from qryTranslatedStrings where (Lang = '" & strLanguage & "')"
    Set rst = dbs.OpenRecordset(strSql, dbOpenDynaset)
    With rst
        Do While Not .EOF
            strKey = BuildKey(Nz(!Context), Nz(!msgid))
            If Not this.dStrings.Exists(strKey) Then this.dStrings.Add strKey, CStr(!ID)
            If Nz(!Translation) <> vbNullString Then
                If Not this.dTranslation.Exists(CStr(!ID)) Then
                    this.dTranslation.Add CStr(!ID), Nz(!Translation)
                End If
            End If
            .MoveNext
        Loop
        .Close
    End With

    ' Reload the ribbon to reflect any translation changes there.
    If RibbonLoaded Then
        ' Export strings to json file for use in the ribbon.
        strPath = GetInstallSettings.strInstallFolder & PathSep & "Ribbon.json"
        WriteFile ConvertToJson(GetRibbonStrings, 2), strPath
        modCOMAddIn.ReloadRibbon
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : CurrentLanguage
' Author    : Adam Waller
' Date      : 5/15/2021
' Purpose   : Return the currently selected language, falling back to operating system
'           : UI language, then to US English.
'---------------------------------------------------------------------------------------
'
Public Function GetCurrentLanguage() As String
    GetCurrentLanguage = Coalesce(this.CurrentLanguage, Me.Language, GetOsLanguage, "en_US")
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetOsLanguage
' Author    : Adam Waller
' Date      : 8/22/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function GetOsLanguage() As String
    GetOsLanguage = LanguageIdToString(LanguageSettings.LanguageID(msoLanguageIDUI))
End Function


'---------------------------------------------------------------------------------------
' Procedure : ImportTranslation
' Author    : Adam Waller
' Date      : 5/15/2021
' Purpose   : Import a translation file. (*.po/*.pot)
'           : See https://www.gnu.org/software/gettext/manual/html_node/PO-Files.html
'           : for details on .po file format.
'---------------------------------------------------------------------------------------
'
Private Sub ImportTranslation(strFile As String)

    Dim strName As String
    Dim strLanguage As String
    Dim strContent As String
    Dim varLines As Variant
    Dim lngStart As Long
    Dim lngLine As Long
    Dim intSection As eFileSection
    Dim intLastSection As eFileSection
    Dim strLine As String
    Dim dItem As Dictionary
    Dim cText As clsConcat
    Dim cHeader As clsConcat
    Dim colEntries As Collection

    ' Read file contents
    strContent = ReadFile(strFile)
    If strContent = vbNullString Then Exit Sub

    ' Add a blank line at the end to make sure we capture the last entry
    strContent = strContent & vbCrLf

    ' Get language from file name
    strName = FSO.GetFileName(strFile)
    If strName = PROJECT_NAME & ".pot" Then
        ' Template file (English strings)
        strLanguage = "en_US"
    Else
        ' Other language
        strLanguage = FSO.GetBaseName(strName)
    End If

    ' Split into lines
    varLines = Split(strContent, vbCrLf)

    ' Initialize structures
    intSection = efsHeader
    intLastSection = efsHeader
    Set colEntries = New Collection
    Set dItem = New Dictionary
    Set cText = New clsConcat
    Set cHeader = New clsConcat
    cText.AppendOnAdd = vbCrLf
    cHeader.AppendOnAdd = vbCrLf

    ' Loop through lines, build collection of entries
    For lngLine = 0 To UBound(varLines)
        strLine = Trim(varLines(lngLine))

        ' Check for new section
        If strLine = vbNullString Then
            intSection = efsUnknown
            lngStart = 1

        ' Check for context entry
        ElseIf StartsWith(strLine, "msgctxt ") Then
            intSection = efsContext
            dItem("Context") = Mid$(strLine, 9)
            lngStart = 9

        ' Check for reference entry
        ElseIf StartsWith(strLine, "#: ") Then
            intSection = efsReference
            dItem("Reference") = Mid$(strLine, 4)
            lngStart = 4

        ' Header
        ElseIf StartsWith(strLine, "# ") Then
            intSection = efsHeader
            lngStart = 3

        ' Blank header line
        ElseIf strLine = "#" Then
            intSection = efsHeader
            lngStart = 2

        ' Message
        ElseIf StartsWith(strLine, "msgid ") Then
            intSection = efsMessageId
            lngStart = 7

        ' Translation
        ElseIf StartsWith(strLine, "msgstr ") Then
            intSection = efsMessageString
            lngStart = 8

        ' Pick back up at first character for multi-line
        Else
            lngStart = 1

        End If

        ' Check for section change
        If intSection <> intLastSection Then
            ' Finish out the last section
            If cText.Length > 0 Then
                cText.Remove Len(vbCrLf)
                Select Case intLastSection
                    Case efsHeader:         dItem("Comments") = UnQ(cText.GetStr)
                    Case efsMessageId:      dItem("msgid") = UnQ(cText.GetStr)
                    Case efsMessageString:  dItem("Translation") = UnQ(cText.GetStr)
                End Select
                cText.Clear
            End If
            ' Add to new section
            cText.Add Mid$(strLine, lngStart)
            intLastSection = intSection
            ' See if we are ready to finish out the previous entry
            If intSection = efsUnknown Then
                If dItem.Count > 1 Then
                    colEntries.Add dItem
                    Set dItem = New Dictionary
                End If
            End If
        Else
            ' Continue current section
            cText.Add Mid$(strLine, lngStart)
        End If

    ' Move to next line in file
    Next lngLine

    ' Export file for debugging parsing process
    'Set dItem = New Dictionary: dItem.Add "Entries", colEntries: WriteFile _
        BuildJsonFile(TypeName(Me), dItem, "Language File Parsed Output"), _
        TranslationsPath & PathSep & strLanguage & ".json"

    ' Save to add-in database
    AddLanguage strLanguage
    If strLanguage = "en_US" Then
        SaveStringEntriesToDatabase colEntries
    Else
        SaveTranslationEntriesToDatabase colEntries, strLanguage
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : SaveStringEntriesToDatabase
' Author    : Adam Waller
' Date      : 11/7/2024
' Purpose   : Save the master list of translation strings (en_US)
'---------------------------------------------------------------------------------------
'
Private Sub SaveStringEntriesToDatabase(colEntries As Collection)

    Dim dbs As Database
    Dim rst As Recordset
    Dim lngCnt As Long
    Dim dItem As Dictionary
    Dim varKey As Variant

    ' This might be the add-in file during development, such as when loading
    ' language strings after a fresh build, or it might be at runtime when
    ' syncronizing translation strings. Either way, we want to run in CodeDB.
    Set dbs = CodeDb

    ' Remove all translations, since we are resetting IDs
    dbs.Execute "delete from tblTranslation", dbFailOnError
    Set this.dStrings = New Dictionary

    ' Reload master list of strings
    dbs.Execute "delete from tblStrings", dbFailOnError
    Set rst = dbs.OpenRecordset("tblStrings", dbOpenDynaset)
    With rst
        For lngCnt = 1 To colEntries.Count
            Set dItem = colEntries(lngCnt)
            ' Build key for entry
            varKey = BuildKey(dItem("Context"), dItem("msgid"))
            ' See if the key already exists for this string:
            If this.dStrings.Exists(varKey) Then
                ' Ignore duplicate entries
            Else
                ' Add string to table
                .AddNew
                    !msgid = dItem("msgid")
                    !Context = ZN(dItem("Context"))
                    !Reference = ZN(dItem("Reference"))
                    !Comments = ZN(dItem("Comments"))
                .Update
                ' Save ID to dictionary using key
                .Bookmark = .LastModified
                this.dStrings.Add varKey, CLng(!ID)
            End If
        Next lngCnt
        .Close
    End With

End Sub


'---------------------------------------------------------------------------------------
' Procedure : SaveTranslationEntriesToDatabase
' Author    : Adam Waller
' Date      : 11/7/2024
' Purpose   : Save translated strings to `tblTranslation`.
'---------------------------------------------------------------------------------------
'
Private Sub SaveTranslationEntriesToDatabase(colEntries As Collection, strLanguage As String)

    Dim dbs As Database
    Dim rst As Recordset
    Dim lngCnt As Long
    Dim dItem As Dictionary
    Dim varKey As Variant
    Dim strID As String
    Dim dAdded As Dictionary

    ' Make sure we don't attempt to add translations before the main strings.
    If this.dStrings Is Nothing Then
        MsgBox2 "Load English Language First", _
            "Please ensure that the default 'en_US' language is loaded before loading other translation files.", _
            "Review language loading logic to correct this issue.", vbExclamation
        Exit Sub
    End If

    ' Reload language
    Set dbs = CodeDb
    dbs.Execute "delete from tblTranslation where Language='" & strLanguage & "'", dbFailOnError
    Set rst = dbs.OpenRecordset("tblTranslation")
    Set this.dTranslation = New Dictionary
    Set dAdded = New Dictionary

    ' Loop through entries, saving unique entries to database
    With rst
        For lngCnt = 1 To colEntries.Count
            Set dItem = colEntries(lngCnt)
            ' Build key for entry
            varKey = BuildKey(dItem("Context"), dItem("msgid"))
            ' See if the key exists for this string:
            If Not this.dStrings.Exists(varKey) Then
                ' Key not found in master list of strings.
                ' Don't load this (likely orphaned) translation string.
            Else
                ' Only add entries that have values
                If Len(dItem("Translation")) Then
                    ' Using a string for the ID so we can use it as a dictionary key
                    strID = this.dStrings(varKey)
                    ' Make sure we have a string ID that has not already been added
                    If Len(strID) And Not dAdded.Exists(strID) Then
                        .AddNew
                            !Language = strLanguage
                            !StringID = CLng(strID)
                            !Translation = dItem("Translation")
                        .Update
                        ' Keep track of which IDs we have added.
                        dAdded.Add strID, vbNullString
                    End If
                End If
            End If
        Next lngCnt
        .Close
    End With

End Sub


'---------------------------------------------------------------------------------------
' Procedure : AddLanguage
' Author    : Adam Waller
' Date      : 5/26/2021
' Purpose   : Add a language entry
'---------------------------------------------------------------------------------------
'
Private Sub AddLanguage(strLanguage As String)

    Dim dbs As Database
    Dim rst As Recordset

    Set dbs = CodeDb

    ' Clear any existing record
    dbs.Execute "delete from tblLanguages where ID='" & strLanguage & "'", dbFailOnError

    ' Add new/replacement record
    Set rst = dbs.OpenRecordset("tblLanguages")
    With rst
        .AddNew
            !ID = strLanguage
            !DisplayName = GetLanguageName(strLanguage)
        .Update
        .Close
    End With

End Sub


'---------------------------------------------------------------------------------------
' Procedure : BuildFileContent
' Author    : Adam Waller
' Date      : 5/19/2021
' Purpose   : Creates the .po/.pot file. (en_US will be treated as a template)
'---------------------------------------------------------------------------------------
'
Private Function BuildFileContent(strLanguage As String) As String

    Const cstrBreak As String = "zzMSAccessVCS"

    Dim dbs As Database
    Dim rst As Recordset

    ' Make sure we are referencing the add-in (code) database, not the open database
    Set dbs = CodeDb

    With New clsConcat
        .AppendOnAdd = vbCrLf

        ' Load strings from database
        ' (Note that sorting does not always seem to be consistent. Tried `Left(key, 255)`
        ' but this still had the same issue with one of the entries going out of order
        ' every other time I run the query.)
        Set rst = dbs.OpenRecordset( _
            "select * from qryTranslatedStrings where Lang='" & strLanguage & "'" & _
            " order by SortRank, key", dbOpenSnapshot)

        ' Loop through strings
        Do While Not rst.EOF
            If rst!msgid = cstrBreak Then Stop
            If Nz(rst!Comments) <> vbNullString Then .Add FormatComment(Nz(rst!Comments))
            If Nz(rst!Reference) <> vbNullString Then .Add "#: ", Nz(rst!Reference)
            If Nz(rst!Context) <> vbNullString Then .Add "msgctxt ", Q(Nz(rst!Context))
            .Add "msgid ", Q(Nz(rst!msgid))
            If strLanguage = "en_TEST" Then
                ' Create test strings by captializing English
                ' (This provides an easy way to verify the translation implementation
                '  in the user interface and log files.)
                .Add "msgstr ", Q(UCase(Nz(rst!msgid)))
            Else
                ' Add the translation string
                .Add "msgstr ", Q(Nz(rst!Translation))
            End If
            .Add vbNullString ' (blank line before new entry)
            rst.MoveNext
        Loop
        rst.Close

        ' Remove final blank line
        If .Length > 2 Then .Remove 2

        ' Return assembled content
        BuildFileContent = .GetStr
    End With

End Function


'---------------------------------------------------------------------------------------
' Procedure : Q
' Author    : Adam Waller
' Date      : 5/19/2021
' Purpose   : Quotes the string, and escapes any embedded quotes. Also breaks long
'           : strings into multiple lines and replaces vbCrLf with \n. Additionally,
'           : \n adds an additional line break.
'---------------------------------------------------------------------------------------
'
Private Function Q(strText As String) As String

    ' Maximum line length (Matches Poedit output)
    Const MAX_LEN As Integer = 74

    Dim strNew As String
    Dim intPos As Integer
    Dim intStart As Integer
    Dim varLines As Variant
    Dim lngLine As Long
    Dim strLine As String

    ' Flag for debugging
    'If StartsWith(strText, "Use this tool to duplicate") Then Stop

    ' Replace newlines and quotes with placeholder
    strNew = Replace(strText, vbCrLf, "\n")
    strNew = Replace(strNew, vbLf, "\n")
    strNew = Replace(strNew, vbCr, "\n")
    strNew = Replace(strNew, """", "\""")

    ' Short curcuit on simple strings
    If InStr(1, strNew, "\n") = 0 And Len(strNew) + 5 < MAX_LEN Then
        ' Return single line
        Q = """" & strNew & """"
        Exit Function
    End If

    ' Split into lines using \n
    varLines = Split(strNew, "\n")

    With New clsConcat

        ' Start with blank string
        .Add """"""

        ' Loop through lines, building return string
        For lngLine = 0 To UBound(varLines)
            strLine = varLines(lngLine)

            ' Add line breaks for over 70 characters.
            ' (80 characters is standard for PO files)
            If Len(strLine) > MAX_LEN Then

                ' Begin at first character
                intStart = 1

                ' Continue while start is less than full length of line
                Do While intStart < Len(strLine)
                    intPos = MAX_LEN
                    ' Look for space closer to max length first
                    intPos = InStrRev(Mid$(strLine, intStart, MAX_LEN), " ", intPos)
                    ' If no space found, use full max length
                    If intPos = 0 Then intPos = MAX_LEN
                    ' Break string here, and move start
                    .Add vbCrLf, """", Mid$(strLine, intStart, intPos), """"
                    intStart = intStart + intPos
                    ' Add final partial string
                    If Len(strLine) - intStart < MAX_LEN Then
                        .Add vbCrLf, """", Mid$(strLine, intStart)
                        Exit Do
                    End If
                Loop
                ' We have now reached the end of that line.
                ' Add split delimiter for all but last line
                If lngLine < UBound(varLines) Then .Add "\n"
                .Add """"
            Else
                ' Skip any final blank line
                If strLine = vbNullString And lngLine = UBound(varLines) Then Exit For
                ' Less than max length. Add full string
                .Add vbCrLf, """", strLine
                ' Add split delimiter for all but last line
                If lngLine < UBound(varLines) Then .Add "\n"
                .Add """"
            End If
        Next lngLine

        ' Return multi-line result
        Q = .GetStr

    End With

End Function


'---------------------------------------------------------------------------------------
' Procedure : UnQ
' Author    : Adam Waller
' Date      : 6/22/2024
' Purpose   : Unquote a string
'---------------------------------------------------------------------------------------
'
Private Function UnQ(strText As String)

    Dim strNew As String

    ' Trim any leading and trailing spaces
    strNew = Trim(strText)

    ' Combine lines that were previously split
    strNew = Replace(strNew, """" & vbCrLf & """", vbNullString)

    ' Restore newlines and quotes
    strNew = Replace(strNew, "\n", vbLf)    ' LF (not CRLF) is standard for gettext
    strNew = Replace(strNew, "\""", """")

    ' Remove beginning and trailing quotes
    If Left$(strNew, 1) = """" Then strNew = Mid$(strNew, 2)
    If Right$(strNew, 1) = """" Then strNew = Left$(strNew, Len(strNew) - 1)

    ' Return result
    UnQ = strNew

End Function


'---------------------------------------------------------------------------------------
' Procedure : FormatComment
' Author    : Adam Waller
' Date      : 11/8/2024
' Purpose   : Format a single or multi-line comment.
'           : Note: This does not automatically wrap lines, but preserves existing
'           : line breaks.
'---------------------------------------------------------------------------------------
'
Private Function FormatComment(strComment As String) As String

    Dim strNew As String
    Dim varLines As Variant
    Dim strLine As String
    Dim lngLine As Long

    ' Standardize on vbCrLf
    strNew = Replace(strComment, vbCrLf, "\n")
    strNew = Replace(strNew, vbCr, "\n")
    strNew = Replace(strNew, vbLf, "\n")
    strNew = Replace(strNew, "\n", vbCrLf)

    With New clsConcat
        ' Split by lines
        varLines = Split(strNew, vbCrLf)
        For lngLine = 0 To UBound(varLines)
            strLine = varLines(lngLine)
            .Add "# ", strLine, vbCrLf
        Next lngLine
        If .Length > 2 Then .Remove Len(vbCrLf)
        FormatComment = .GetStr
    End With

End Function


'---------------------------------------------------------------------------------------
' Procedure : LanguageIdToString
' Author    : Adam Waller
' Date      : 8/21/2024
' Purpose   : Used to translate the UI language constant to the language code.
'           : Application.LanguageSettings.LanguageID(msoLanguageIDUI)
'           : Compiled using vlookup in Excel from the following two sources:
'           : https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-lcid/
'           : https://learn.microsoft.com/en-us/office/vba/api/Office.MsoLanguageID
'---------------------------------------------------------------------------------------
'
Private Function LanguageIdToString(lngID) As String

    Select Case lngID
        Case 1078:  LanguageIdToString = "af_ZA" 'South Africa (msoLanguageIDAfrikaans)
        Case 1052:  LanguageIdToString = "sq_AL" 'Albania (msoLanguageIDAlbanian)
        Case 1118:  LanguageIdToString = "am_ET" 'Ethiopia (msoLanguageIDAmharic)
        Case 1025:  LanguageIdToString = "ar_SA" 'Saudi Arabia (msoLanguageIDArabic)
        Case 5121:  LanguageIdToString = "ar_DZ" 'Algeria (msoLanguageIDArabicAlgeria)
        Case 15361: LanguageIdToString = "ar_BH" 'Bahrain (msoLanguageIDArabicBahrain)
        Case 3073:  LanguageIdToString = "ar_EG" 'Egypt (msoLanguageIDArabicEgypt)
        Case 2049:  LanguageIdToString = "ar_IQ" 'Iraq (msoLanguageIDArabicIraq)
        Case 11265: LanguageIdToString = "ar_JO" 'Jordan (msoLanguageIDArabicJordan)
        Case 13313: LanguageIdToString = "ar_KW" 'Kuwait (msoLanguageIDArabicKuwait)
        Case 12289: LanguageIdToString = "ar_LB" 'Lebanon (msoLanguageIDArabicLebanon)
        Case 4097:  LanguageIdToString = "ar_LY" 'Libya (msoLanguageIDArabicLibya)
        Case 6145:  LanguageIdToString = "ar_MA" 'Morocco (msoLanguageIDArabicMorocco)
        Case 8193:  LanguageIdToString = "ar_OM" 'Oman (msoLanguageIDArabicOman)
        Case 16385: LanguageIdToString = "ar_QA" 'Qatar (msoLanguageIDArabicQatar)
        Case 10241: LanguageIdToString = "ar_SY" 'Syria (msoLanguageIDArabicSyria)
        Case 7169:  LanguageIdToString = "ar_TN" 'Tunisia (msoLanguageIDArabicTunisia)
        Case 14337: LanguageIdToString = "ar_AE" 'U.A.E. (msoLanguageIDArabicUAE)
        Case 9217:  LanguageIdToString = "ar_YE" 'Yemen (msoLanguageIDArabicYemen)
        Case 1067:  LanguageIdToString = "hy_AM" 'Armenia (msoLanguageIDArmenian)
        Case 1101:  LanguageIdToString = "as_IN" 'India (msoLanguageIDAssamese)
        Case 2092:  LanguageIdToString = "az_Cyrl_AZ" 'Azerbaijan (msoLanguageIDAzeriCyrillic)
        Case 1068:  LanguageIdToString = "az_Latn_AZ" 'Azerbaijan (msoLanguageIDAzeriLatin)
        Case 1069:  LanguageIdToString = "eu_ES" 'Spain (msoLanguageIDBasque)
        Case 2067:  LanguageIdToString = "nl_BE" 'Belgium (msoLanguageIDBelgianDutch)
        Case 2060:  LanguageIdToString = "fr_BE" 'Belgium (msoLanguageIDBelgianFrench)
        Case 1093:  LanguageIdToString = "bn_IN" 'India (msoLanguageIDBengali)
        Case 4122:  LanguageIdToString = "hr_BA" 'Bosnia and Herzegovina (msoLanguageIDBosnian)
        Case 8218:  LanguageIdToString = "bs_Cyrl_BA" 'Bosnia and Herzegovina (msoLanguageIDBosnianBosniaHerzegovinaCyrillic)
        Case 5146:  LanguageIdToString = "bs_Latn_BA" 'Bosnia and Herzegovina (msoLanguageIDBosnianBosniaHerzegovinaLatin)
        Case 1046:  LanguageIdToString = "pt_BR" 'Brazil (msoLanguageIDBrazilianPortuguese)
        Case 1026:  LanguageIdToString = "bg_BG" 'Bulgaria (msoLanguageIDBulgarian)
        Case 1109:  LanguageIdToString = "my_MM" 'Myanmar (msoLanguageIDBurmese)
        Case 1059:  LanguageIdToString = "be_BY" 'Belarus (msoLanguageIDByelorussian)
        Case 1027:  LanguageIdToString = "ca_ES" 'Spain (msoLanguageIDCatalan)
        Case 1116:  LanguageIdToString = "chr_Cher_US" 'United States (msoLanguageIDCherokee)
        Case 3076:  LanguageIdToString = "zh_HK" 'Hong Kong S.A.R. (msoLanguageIDChineseHongKongSAR)
        Case 5124:  LanguageIdToString = "zh_MO" 'Macao S.A.R. (msoLanguageIDChineseMacaoSAR)
        Case 4100:  LanguageIdToString = "zh_SG" 'Singapore (msoLanguageIDChineseSingapore)
        Case 1050:  LanguageIdToString = "hr_HR" 'Croatia (msoLanguageIDCroatian)
        Case 1029:  LanguageIdToString = "cs_CZ" 'Czech Republic (msoLanguageIDCzech)
        Case 1030:  LanguageIdToString = "da_DK" 'Denmark (msoLanguageIDDanish)
        Case 1125:  LanguageIdToString = "dv_MV" 'Maldives (msoLanguageIDDivehi)
        Case 1043:  LanguageIdToString = "nl_NL" 'Netherlands (msoLanguageIDDutch)
        Case 3081:  LanguageIdToString = "en_AU" 'Australia (msoLanguageIDEnglishAUS)
        Case 10249: LanguageIdToString = "en_BZ" 'Belize (msoLanguageIDEnglishBelize)
        Case 4105:  LanguageIdToString = "en_CA" 'Canada (msoLanguageIDEnglishCanadian)
        Case 9225:  LanguageIdToString = "en_029" 'Caribbean (msoLanguageIDEnglishCaribbean)
        Case 6153:  LanguageIdToString = "en_IE" 'Ireland (msoLanguageIDEnglishIreland)
        Case 8201:  LanguageIdToString = "en_JM" 'Jamaica (msoLanguageIDEnglishJamaica)
        Case 5129:  LanguageIdToString = "en_NZ" 'New Zealand (msoLanguageIDEnglishNewZealand)
        Case 13321: LanguageIdToString = "en_PH" 'Republic of the Philippines (msoLanguageIDEnglishPhilippines)
        Case 7177:  LanguageIdToString = "en_ZA" 'South Africa (msoLanguageIDEnglishSouthAfrica)
        Case 11273: LanguageIdToString = "en_TT" 'Trinidad and Tobago (msoLanguageIDEnglishTrinidadTobago)
        Case 2057:  LanguageIdToString = "en_GB" 'United Kingdom (msoLanguageIDEnglishUK)
        Case 1033:  LanguageIdToString = "en_US" 'United States (msoLanguageIDEnglishUS)
        Case 12297: LanguageIdToString = "en_ZW" 'Zimbabwe (msoLanguageIDEnglishZimbabwe)
        Case 1061:  LanguageIdToString = "et_EE" 'Estonia (msoLanguageIDEstonian)
        Case 1080:  LanguageIdToString = "fo_FO" 'Faroe Islands (msoLanguageIDFaeroese)
        Case 1065:  LanguageIdToString = "fa_IR" 'Iran (msoLanguageIDFarsi)
        Case 1124:  LanguageIdToString = "fil_PH" 'Philippines (msoLanguageIDFilipino)
        Case 1035:  LanguageIdToString = "fi_FI" 'Finland (msoLanguageIDFinnish)
        Case 1036:  LanguageIdToString = "fr_FR" 'France (msoLanguageIDFrench)
        Case 11276: LanguageIdToString = "fr_CM" 'Cameroon (msoLanguageIDFrenchCameroon)
        Case 3084:  LanguageIdToString = "fr_CA" 'Canada (msoLanguageIDFrenchCanadian)
        Case 9228:  LanguageIdToString = "fr_CD" 'Congo, DRC (msoLanguageIDFrenchCongoDRC)
        Case 12300: LanguageIdToString = "fr_CI" 'Côte d'Ivoire (msoLanguageIDFrenchCotedIvoire)
        Case 15372: LanguageIdToString = "fr_HT" 'Haiti (msoLanguageIDFrenchHaiti)
        Case 5132:  LanguageIdToString = "fr_LU" 'Luxembourg (msoLanguageIDFrenchLuxembourg)
        Case 13324: LanguageIdToString = "fr_ML" 'Mali (msoLanguageIDFrenchMali)
        Case 6156:  LanguageIdToString = "fr_MC" 'Principality of Monaco (msoLanguageIDFrenchMonaco)
        Case 14348: LanguageIdToString = "fr_MA" 'Morocco (msoLanguageIDFrenchMorocco)
        Case 8204:  LanguageIdToString = "fr_RE" 'Reunion (msoLanguageIDFrenchReunion)
        Case 10252: LanguageIdToString = "fr_SN" 'Senegal (msoLanguageIDFrenchSenegal)
        Case 7180:  LanguageIdToString = "fr_029" 'Caribbean (msoLanguageIDFrenchWestIndies)
        Case 1122:  LanguageIdToString = "fy_NL" 'Netherlands (msoLanguageIDFrisianNetherlands)
        Case 1127:  LanguageIdToString = "ff_NG" 'Nigeria (msoLanguageIDFulfulde)
        Case 1127:  LanguageIdToString = "ff_Latn_NG" 'Nigeria (msoLanguageIDFulfulde)
        Case 2108:  LanguageIdToString = "ga_IE" 'Ireland (msoLanguageIDGaelicIreland)
        Case 1110:  LanguageIdToString = "gl_ES" 'Spain (msoLanguageIDGalician)
        Case 1079:  LanguageIdToString = "ka_GE" 'Georgia (msoLanguageIDGeorgian)
        Case 1031:  LanguageIdToString = "de_DE" 'Germany (msoLanguageIDGerman)
        Case 3079:  LanguageIdToString = "de_AT" 'Austria (msoLanguageIDGermanAustria)
        Case 5127:  LanguageIdToString = "de_LI" 'Liechtenstein (msoLanguageIDGermanLiechtenstein)
        Case 4103:  LanguageIdToString = "de_LU" 'Luxembourg (msoLanguageIDGermanLuxembourg)
        Case 1032:  LanguageIdToString = "el_GR" 'Greece (msoLanguageIDGreek)
        Case 1140:  LanguageIdToString = "gn_PY" 'Paraguay (msoLanguageIDGuarani)
        Case 1095:  LanguageIdToString = "gu_IN" 'India (msoLanguageIDGujarati)
        Case 1128:  LanguageIdToString = "ha_Latn_NG" 'Nigeria (msoLanguageIDHausa)
        Case 1141:  LanguageIdToString = "haw_US" 'United States (msoLanguageIDHawaiian)
        Case 1037:  LanguageIdToString = "he_IL" 'Israel (msoLanguageIDHebrew)
        Case 1081:  LanguageIdToString = "hi_IN" 'India (msoLanguageIDHindi)
        Case 1038:  LanguageIdToString = "hu_HU" 'Hungary (msoLanguageIDHungarian)
        Case 1039:  LanguageIdToString = "is_IS" 'Iceland (msoLanguageIDIcelandic)
        Case 1136:  LanguageIdToString = "ig_NG" 'Nigeria (msoLanguageIDIgbo)
        Case 1057:  LanguageIdToString = "id_ID" 'Indonesia (msoLanguageIDIndonesian)
        Case 1117:  LanguageIdToString = "iu_Cans_CA" 'Canada (msoLanguageIDInuktitut)
        Case 1040:  LanguageIdToString = "it_IT" 'Italy (msoLanguageIDItalian)
        Case 1041:  LanguageIdToString = "ja_JP" 'Japan (msoLanguageIDJapanese)
        Case 1099:  LanguageIdToString = "kn_IN" 'India (msoLanguageIDKannada)
        Case 1137:  LanguageIdToString = "kr_Latn_NG" 'Nigeria (msoLanguageIDKanuri)
        Case 1120:  LanguageIdToString = "ks_Arab" 'Perso-Arabic (msoLanguageIDKashmiri)
        Case 2144:  LanguageIdToString = "ks_Deva_IN" 'India (msoLanguageIDKashmiriDevanagari)
        Case 1087:  LanguageIdToString = "kk_KZ" 'Kazakhstan (msoLanguageIDKazakh)
        Case 1107:  LanguageIdToString = "km_KH" 'Cambodia (msoLanguageIDKhmer)
        Case 1088:  LanguageIdToString = "ky_KG" 'Kyrgyzstan (msoLanguageIDKirghiz)
        Case 1111:  LanguageIdToString = "kok_IN" 'India (msoLanguageIDKonkani)
        Case 1042:  LanguageIdToString = "ko_KR" 'Korea (msoLanguageIDKorean)
        Case 1108:  LanguageIdToString = "lo_LA" 'Lao P.D.R. (msoLanguageIDLao)
        Case 1142:  LanguageIdToString = "la_VA" 'Vatican City (msoLanguageIDLatin)
        Case 1062:  LanguageIdToString = "lv_LV" 'Latvia (msoLanguageIDLatvian)
        Case 1063:  LanguageIdToString = "lt_LT" 'Lithuania (msoLanguageIDLithuanian)
        Case 1071:  LanguageIdToString = "mk_MK" 'North Macedonia (msoLanguageIDMacedonianFYROM)
        Case 1100:  LanguageIdToString = "ml_IN" 'India (msoLanguageIDMalayalam)
        Case 2110:  LanguageIdToString = "ms_BN" 'Brunei Darussalam (msoLanguageIDMalayBruneiDarussalam)
        Case 1086:  LanguageIdToString = "ms_MY" 'Malaysia (msoLanguageIDMalaysian)
        Case 1082:  LanguageIdToString = "mt_MT" 'Malta (msoLanguageIDMaltese)
        Case 1153:  LanguageIdToString = "mi_NZ" 'New Zealand (msoLanguageIDMaori)
        Case 1102:  LanguageIdToString = "mr_IN" 'India (msoLanguageIDMarathi)
        Case 2058:  LanguageIdToString = "es_MX" 'Mexico (msoLanguageIDMexicanSpanish)
        Case 1104:  LanguageIdToString = "mn_MN" 'Mongolia (msoLanguageIDMongolian)
        Case 1121:  LanguageIdToString = "ne_NP" 'Nepal (msoLanguageIDNepali)
        Case 1044:  LanguageIdToString = "nb_NO" 'Norway (msoLanguageIDNorwegianBokmol)
        Case 2068:  LanguageIdToString = "nn_NO" 'Norway (msoLanguageIDNorwegianNynorsk)
        Case 1096:  LanguageIdToString = "or_IN" 'India (msoLanguageIDOriya)
        Case 1138:  LanguageIdToString = "om_ET" 'Ethiopia (msoLanguageIDOromo)
        Case 1123:  LanguageIdToString = "ps_AF" 'Afghanistan (msoLanguageIDPashto)
        Case 1045:  LanguageIdToString = "pl_PL" 'Poland (msoLanguageIDPolish)
        Case 2070:  LanguageIdToString = "pt_PT" 'Portugal (msoLanguageIDPortuguese)
        Case 1094:  LanguageIdToString = "pa_IN" 'India (msoLanguageIDPunjabi)
        Case 1131:  LanguageIdToString = "quz_BO" 'Bolivia (msoLanguageIDQuechuaBolivia)
        Case 2155:  LanguageIdToString = "quz_EC" 'Ecuador (msoLanguageIDQuechuaEcuador)
        Case 3179:  LanguageIdToString = "quz_PE" 'Peru (msoLanguageIDQuechuaPeru)
        Case 1047:  LanguageIdToString = "rm_CH" 'Switzerland (msoLanguageIDRhaetoRomanic)
        Case 1048:  LanguageIdToString = "ro_RO" 'Romania (msoLanguageIDRomanian)
        Case 2072:  LanguageIdToString = "ro_MD" 'Moldova (msoLanguageIDRomanianMoldova)
        Case 1049:  LanguageIdToString = "ru_RU" 'Russia (msoLanguageIDRussian)
        Case 2073:  LanguageIdToString = "ru_MD" 'Moldova (msoLanguageIDRussianMoldova)
        Case 1083:  LanguageIdToString = "se_NO" 'Norway (msoLanguageIDSamiLappish)
        Case 1103:  LanguageIdToString = "sa_IN" 'India (msoLanguageIDSanskrit)
        Case 1132:  LanguageIdToString = "nso_ZA" 'South Africa (msoLanguageIDSepedi)
        Case 7194:  LanguageIdToString = "sr_Cyrl_BA" 'Bosnia and Herzegovina (msoLanguageIDSerbianBosniaHerzegovinaCyrillic)
        Case 6170:  LanguageIdToString = "sr_Latn_BA" 'Bosnia and Herzegovina (msoLanguageIDSerbianBosniaHerzegovinaLatin)
        Case 3098:  LanguageIdToString = "sr_Cyrl_CS" 'Serbia and Montenegro (Former) (msoLanguageIDSerbianCyrillic)
        Case 2074:  LanguageIdToString = "sr_Latn_CS" 'Serbia and Montenegro (Former) (msoLanguageIDSerbianLatin)
        Case 1072:  LanguageIdToString = "st_ZA" 'South Africa (msoLanguageIDSesotho)
        Case 2052:  LanguageIdToString = "zh_CN" 'People's Republic of China (msoLanguageIDSimplifiedChinese)
        Case 2137:  LanguageIdToString = "sd_Arab_PK" 'Islamic Republic of Pakistan (msoLanguageIDSindhiPakistan)
        Case 1115:  LanguageIdToString = "si_LK" 'Sri Lanka (msoLanguageIDSinhalese)
        Case 1051:  LanguageIdToString = "sk_SK" 'Slovakia (msoLanguageIDSlovak)
        Case 1060:  LanguageIdToString = "sl_SI" 'Slovenia (msoLanguageIDSlovenian)
        Case 1143:  LanguageIdToString = "so_SO" 'Somalia (msoLanguageIDSomali)
        Case 1070:  LanguageIdToString = "hsb_DE" 'Germany (msoLanguageIDSorbian)
        Case 1034:  LanguageIdToString = "es_ES_tradnl" 'Spain (msoLanguageIDSpanish)
        Case 11274: LanguageIdToString = "es_AR" 'Argentina (msoLanguageIDSpanishArgentina)
        Case 16394: LanguageIdToString = "es_BO" 'Bolivia (msoLanguageIDSpanishBolivia)
        Case 13322: LanguageIdToString = "es_CL" 'Chile (msoLanguageIDSpanishChile)
        Case 9226:  LanguageIdToString = "es_CO" 'Colombia (msoLanguageIDSpanishColombia)
        Case 5130:  LanguageIdToString = "es_CR" 'Costa Rica (msoLanguageIDSpanishCostaRica)
        Case 7178:  LanguageIdToString = "es_DO" 'Dominican Republic (msoLanguageIDSpanishDominicanRepublic)
        Case 12298: LanguageIdToString = "es_EC" 'Ecuador (msoLanguageIDSpanishEcuador)
        Case 17418: LanguageIdToString = "es_SV" 'El Salvador (msoLanguageIDSpanishElSalvador)
        Case 4106:  LanguageIdToString = "es_GT" 'Guatemala (msoLanguageIDSpanishGuatemala)
        Case 18442: LanguageIdToString = "es_HN" 'Honduras (msoLanguageIDSpanishHonduras)
        Case 3082:  LanguageIdToString = "es_ES" 'Spain (msoLanguageIDSpanishModernSort)
        Case 19466: LanguageIdToString = "es_NI" 'Nicaragua (msoLanguageIDSpanishNicaragua)
        Case 6154:  LanguageIdToString = "es_PA" 'Panama (msoLanguageIDSpanishPanama)
        Case 15370: LanguageIdToString = "es_PY" 'Paraguay (msoLanguageIDSpanishParaguay)
        Case 10250: LanguageIdToString = "es_PE" 'Peru (msoLanguageIDSpanishPeru)
        Case 20490: LanguageIdToString = "es_PR" 'Puerto Rico (msoLanguageIDSpanishPuertoRico)
        Case 14346: LanguageIdToString = "es_UY" 'Uruguay (msoLanguageIDSpanishUruguay)
        Case 8202:  LanguageIdToString = "es_VE" 'Bolivarian Republic of Venezuela (msoLanguageIDSpanishVenezuela)
        Case 1089:  LanguageIdToString = "sw_KE" 'Kenya (msoLanguageIDSwahili)
        Case 1053:  LanguageIdToString = "sv_SE" 'Sweden (msoLanguageIDSwedish)
        Case 2077:  LanguageIdToString = "sv_FI" 'Finland (msoLanguageIDSwedishFinland)
        Case 4108:  LanguageIdToString = "fr_CH" 'Switzerland (msoLanguageIDSwissFrench)
        Case 2055:  LanguageIdToString = "de_CH" 'Switzerland (msoLanguageIDSwissGerman)
        Case 2064:  LanguageIdToString = "it_CH" 'Switzerland (msoLanguageIDSwissItalian)
        Case 1114:  LanguageIdToString = "syr_SY" 'Syria (msoLanguageIDSyriac)
        Case 1064:  LanguageIdToString = "tg_Cyrl_TJ" 'Tajikistan (msoLanguageIDTajik)
        Case 1119:  LanguageIdToString = "tzm_Arab_MA" 'Morocco (msoLanguageIDTamazight)
        Case 2143:  LanguageIdToString = "tzm_Latn_DZ" 'Algeria (msoLanguageIDTamazightLatin)
        Case 1097:  LanguageIdToString = "ta_IN" 'India (msoLanguageIDTamil)
        Case 1092:  LanguageIdToString = "tt_RU" 'Russia (msoLanguageIDTatar)
        Case 1098:  LanguageIdToString = "te_IN" 'India (msoLanguageIDTelugu)
        Case 1054:  LanguageIdToString = "th_TH" 'Thailand (msoLanguageIDThai)
        Case 1105:  LanguageIdToString = "bo_CN" 'People's Republic of China (msoLanguageIDTibetan)
        Case 2163:  LanguageIdToString = "ti_ER" 'Eritrea (msoLanguageIDTigrignaEritrea)
        Case 1139:  LanguageIdToString = "ti_ET" 'Ethiopia (msoLanguageIDTigrignaEthiopic)
        Case 1028:  LanguageIdToString = "zh_TW" 'Taiwan (msoLanguageIDTraditionalChinese)
        Case 1073:  LanguageIdToString = "ts_ZA" 'South Africa (msoLanguageIDTsonga)
        Case 1074:  LanguageIdToString = "tn_ZA" 'South Africa (msoLanguageIDTswana)
        Case 1055:  LanguageIdToString = "tr_TR" 'Turkey (msoLanguageIDTurkish)
        Case 1090:  LanguageIdToString = "tk_TM" 'Turkmenistan (msoLanguageIDTurkmen)
        Case 1058:  LanguageIdToString = "uk_UA" 'Ukraine (msoLanguageIDUkrainian)
        Case 1056:  LanguageIdToString = "ur_PK" 'Islamic Republic of Pakistan (msoLanguageIDUrdu)
        Case 2115:  LanguageIdToString = "uz_Cyrl_UZ" 'Uzbekistan (msoLanguageIDUzbekCyrillic)
        Case 1091:  LanguageIdToString = "uz_Latn_UZ" 'Uzbekistan (msoLanguageIDUzbekLatin)
        Case 1075:  LanguageIdToString = "ve_ZA" 'South Africa (msoLanguageIDVenda)
        Case 1066:  LanguageIdToString = "vi_VN" 'Vietnam (msoLanguageIDVietnamese)
        Case 1106:  LanguageIdToString = "cy_GB" 'United Kingdom (msoLanguageIDWelsh)
        Case 1076:  LanguageIdToString = "xh_ZA" 'South Africa (msoLanguageIDXhosa)
        Case 1144:  LanguageIdToString = "ii_CN" 'People's Republic of China (msoLanguageIDYi)
        Case 1085:  LanguageIdToString = "yi_001" 'World (msoLanguageIDYiddish)
        Case 1130:  LanguageIdToString = "yo_NG" 'Nigeria (msoLanguageIDYoruba)
        Case 1077:  LanguageIdToString = "zu_ZA" 'South Africa (msoLanguageIDZulu)
    End Select

End Function


'---------------------------------------------------------------------------------------
' Procedure : Class_Terminate
' Author    : Adam Waller
' Date      : 8/23/2024
' Purpose   : Clear object references
'---------------------------------------------------------------------------------------
'
Private Sub Class_Terminate()
    Set this.dStrings = Nothing
    Set this.dTranslation = Nothing
    Set this.dTranslated = Nothing
End Sub
