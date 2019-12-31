Option Compare Database
Option Explicit
Option Private Module

Private Const cstrSpacer As String = "-------------------------------"

' Keep a persistent reference to file system object after initializing version control.
' This way we don't have to recreate this object dozens of times while using VCS.
Private m_FSO As Scripting.FileSystemObject

'---------------------------------------------------------------------------------------
' Procedure : ExportAllSource
' Author    : Adam Waller
' Date      : 1/18/2019
' Purpose   : Exports all source files for the current project.
'---------------------------------------------------------------------------------------
Public Sub ExportAllSource(cModel As IVersionControl)
    
    Dim dbs As DAO.Database
    Dim strSourcePath As String
    Dim strObjectPath As String
    Dim qry As Object
    Dim doc As Object
    Dim strLabel As String
    Dim strType As String
    Dim intType As AcObjectType
    Dim intObjCnt As Integer
    Dim intObjDataCnt As Integer
    Dim objContainer As Object
    Dim sngStart As Single
    Dim strName As String
    Dim colADO As New Collection
    Dim colContainers As New Collection
    Dim varType As Variant
    Dim strData As String
    Dim blnSkipFile As Boolean
    Dim strFile As String
    Dim dteLastCompact As Date
    Dim dteModified As Date

    ' Option used with fast saves
    If cModel.FastSave Then
        strData = GetDBProperty("InitiatedCompactRepair")
        If IsDate(strData) Then dteLastCompact = CDate(strData)
    End If
    
    Set dbs = CurrentDb
    sngStart = Timer
    Set colVerifiedPaths = New Collection   ' Reset cache

    With cModel
        .Log cstrSpacer
        .Log "Beginning Export of all Source", False
        .Log CurrentProject.Name
        .Log "VCS Version " & GetVCSVersion
        If .FastSave Then .Log "Using Fast Save"
        .Log Now()
    End With
    
    ' Read in options from model
    strSourcePath = cModel.ExportBaseFolder

    ' Make sure we have a path for the source files
    VerifyPath strSourcePath

    ' Display header in debug output
    Debug.Print
    Debug.Print cstrSpacer
    Debug.Print "  Exporting All Source"
    Debug.Print "  Export Path: " & strSourcePath
    Debug.Print cstrSpacer

    ' Process queries
    
    If CurrentProject.ProjectType = acMDB Then
        ' Standard Access Project
        strObjectPath = strSourcePath & "queries\"
        ClearOrphanedSourceFiles strObjectPath, dbs.QueryDefs, cModel, "bas", "sql"
        cModel.Log cstrSpacer, cModel.ShowDebug
        cModel.Log PadRight("Exporting queries...", 24), True, cModel.ShowDebug
        cModel.Log "", cModel.ShowDebug
        intObjCnt = 0
        For Each qry In dbs.QueryDefs
            If Left(qry.Name, 1) <> "~" Then
                strFile = strObjectPath & GetSafeFileName(qry.Name) & ".bas"
                ExportObject acQuery, qry.Name, strFile, cModel
                intObjCnt = intObjCnt + 1
            End If
        Next
        If cModel.ShowDebug Then
            cModel.Log "[" & intObjCnt & "] queries exported."
        Else
            cModel.Log "[" & intObjCnt & "]"
        End If
    Else
        ' ADP project (Several types of 'queries' involved)
        With colADO
            .Add Array("views", "sql", CurrentData.AllViews)
            .Add Array("functions", "sql", CurrentData.AllFunctions)
            .Add Array("procedures", "sql", CurrentData.AllStoredProcedures)
            .Add Array("tables", "tdf", CurrentData.AllTables)
            '.Add Array("diagrams", CurrentData.AllDatabaseDiagrams) ' (Not supported in Access 2010)
            '.Add Array("queries", CurrentData.AllQueries) ' (Combination of views, functions and proceedures)
        End With
        
        ' Clear any triggers if the triggers folder exists.
        If FSO.FolderExists(strObjectPath & "triggers\") Then
            If Not cModel.FastSave Then ClearTextFilesFromDir strObjectPath & "triggers\", "sql"
        End If
        
        ' Process triggers
        cModel.Log cstrSpacer, cModel.ShowDebug
        cModel.Log PadRight("Exporting triggers...", 24), True, cModel.ShowDebug
        cModel.Log "", cModel.ShowDebug
        ExportADPTriggers cModel, strSourcePath & "triggers\"
        
        ' Loop through each type, exporting SQL definitions
        For Each varType In colADO
            strObjectPath = strSourcePath & varType(0) & "\"
            VerifyPath strObjectPath
            
            
            ''''' Wait to clear tables (or other objects) since we need to check the modified date of the file.
            Set objContainer = varType(2)
            ClearOrphanedSourceFiles strObjectPath, objContainer, cModel, varType(1)
            
            cModel.Log cstrSpacer, cModel.ShowDebug
            cModel.Log PadRight("Exporting " & varType(0) & "...", 24), , cModel.ShowDebug
            cModel.Log "", cModel.ShowDebug
            intObjCnt = 0
            For Each qry In varType(2)
                blnSkipFile = False
                strFile = strObjectPath & GetSafeFileName(StripDboPrefix(qry.Name)) & "." & varType(1)
                ' Fast save options
                If cModel.FastSave Then
                    dteModified = GetSQLObjectModifiedDate(qry.Name, varType(0))
                    'dteModified = #1/1/2000#
                    If FSO.FileExists(strFile) Then
                        If dteModified < FileDateTime(strFile) Then
                            ' Object does not appear to have been modified.
                            blnSkipFile = True
                        End If
                    End If
                End If
                If Not blnSkipFile Then
                    If varType(0) = "tables" Then
                        strData = GetADPTableDef(qry.Name)
                    Else
                        strData = GetSQLObjectDefinitionForADP(qry.Name)
                    End If
                End If

                If blnSkipFile Then
                    cModel.Log "  (Skipping '" & qry.Name & "')", cModel.ShowDebug
                Else
                    WriteFile strData, strFile
                    cModel.Log "  " & qry.Name, cModel.ShowDebug
                End If
                intObjCnt = intObjCnt + 1
                ' Check for table/query data export
                If InCollection(cModel.TablesToSaveData, qry.Name) Then
                    DoCmd.OutputTo acOutputServerView, qry.Name, acFormatTXT, strObjectPath & GetSafeFileName(StripDboPrefix(qry.Name)) & ".txt", False
                    cModel.Log "    Data exported", cModel.ShowDebug
                End If
            Next qry
            If cModel.ShowDebug Then
                cModel.Log "[" & intObjCnt & "] " & varType(0) & " exported."
            Else
                cModel.Log "[" & intObjCnt & "]"
            End If
        Next varType
    End If

    ' Clear the cached variables
    GetSQLObjectModifiedDate "", ""
    
    ' Get the forms, reports, macros, and modules
    Set colContainers = New Collection
    With colContainers
        .Add Array("forms", CurrentProject.AllForms, acForm)
        .Add Array("reports", CurrentProject.AllReports, acReport)
        .Add Array("macros", CurrentProject.AllMacros, acMacro)
        .Add Array("modules", CurrentProject.AllModules, acModule)
    End With
    
    ' Loop through main database objects
    For Each varType In colContainers

        strLabel = varType(0)
        Set objContainer = varType(1)
        intType = varType(2)
        strObjectPath = strSourcePath & strLabel & "\"
        intObjCnt = 0
    
        ' Clear out any orphaned source files
        ClearOrphanedSourceFiles strObjectPath, objContainer, cModel, "bas", "pv"
        
        ' Show progress
        cModel.Log cstrSpacer, cModel.ShowDebug
        cModel.Log PadRight("Exporting " & strLabel & "...", 24), , cModel.ShowDebug
        cModel.Log "", cModel.ShowDebug
        
        ' Loop through objects in container
        For Each doc In objContainer
            If (Left(doc.Name, 1) <> "~") Then
                ' Get file name (without extension)
                strFile = strObjectPath & StripDboPrefix(GetSafeFileName(doc.Name))
                ExportObject intType, doc.Name, strFile & ".bas", cModel
                If intType = acReport Then
                    If cModel.SavePrintVars Then ExportPrintVars doc.Name, strFile & ".pv", cModel
                End If
                intObjCnt = intObjCnt + 1
            End If
        Next
        
        ' Show total number of objects
        If cModel.ShowDebug Then
            cModel.Log "[" & intObjCnt & "] " & strLabel & " exported."
        Else
            cModel.Log "[" & intObjCnt & "]"
        End If

    Next varType

    ' Export references
    cModel.Log cstrSpacer, cModel.ShowDebug
    cModel.Log PadRight("Exporting references...", 24), , cModel.ShowDebug
    cModel.Log "", cModel.ShowDebug
    ExportReferences strSourcePath, cModel
    
    ' Export database properties
    cModel.Log cstrSpacer, cModel.ShowDebug
    cModel.Log PadRight("Exporting properties...", 24), , cModel.ShowDebug
    cModel.Log "", cModel.ShowDebug
    ExportProperties strSourcePath, cModel
    
    ' Export Import/Export Specifications
    cModel.Log cstrSpacer, cModel.ShowDebug
    cModel.Log PadRight("Exporting specs...", 24), , cModel.ShowDebug
    cModel.Log "", cModel.ShowDebug
    ExportSpecs strSourcePath, cModel
    
    

'-------------------------mdb table export------------------------
    
    If CurrentProject.ProjectType = acMDB Then
                
        Dim td As TableDef
        Dim tds As TableDefs
        Set tds = dbs.TableDefs
    
        If cModel.TablesToSaveData.Count = 0 Then
            strObjectPath = strSourcePath & "tables"
            If FSO.FolderExists(strObjectPath) Then ClearOrphanedSourceFiles strObjectPath & "\", Nothing, cModel, "txt"
        Else
            ' Only create this folder if we are actually saving table data
            MkDirIfNotExist strSourcePath & "tables\"
            ClearOrphanedSourceFiles strSourcePath & "tables\", dbs.TableDefs, cModel, "txt"
        End If
        
        strLabel = "tbldef"
        strType = "Table_Def"
        intType = acTable
        strObjectPath = strSourcePath & "tbldefs\"
        intObjCnt = 0
        intObjDataCnt = 0
        
        
        ' Verify path and clear any existing files
        VerifyPath Left(strObjectPath, InStrRev(strObjectPath, "\"))
        ClearOrphanedSourceFiles strObjectPath, tds, cModel, "LNKD", "sql", "xml", "bas"

        cModel.Log cstrSpacer, cModel.ShowDebug
        cModel.Log PadRight("Exporting " & strLabel & "...", 24), , cModel.ShowDebug
        cModel.Log "", cModel.ShowDebug
        
        For Each td In tds
            ' This is not a system table
            ' this is not a temporary table
            If Left$(td.Name, 4) <> "MSys" And _
                Left$(td.Name, 1) <> "~" Then
                If Len(td.connect) = 0 Then ' this is not an external table
                    ExportTableDef td.Name, strObjectPath, cModel
                    If InCollection(cModel.TablesToSaveData, "*") Then
                        ExportTableData CStr(td.Name), strSourcePath & "tables\", cModel
                        If Len(Dir(strSourcePath & "tables\" & td.Name & ".txt")) > 0 Then
                            intObjDataCnt = intObjDataCnt + 1
                        End If
                    ElseIf InCollection(cModel.TablesToSaveData, td.Name) Then
                        modTable.ExportTableData CStr(td.Name), strSourcePath & "tables\", cModel
                        intObjDataCnt = intObjDataCnt + 1
                    'else don't export table data
                    End If
    
                Else
                    modTable.ExportLinkedTable td.Name, strObjectPath, cModel
                End If
                
                intObjCnt = intObjCnt + 1
                
            End If
        Next
        
        If cModel.ShowDebug Then
            cModel.Log "[" & intObjCnt & "] tbldefs exported."
        Else
            cModel.Log "[" & intObjCnt & "]"
        End If
    
        ' Export relationships (MDB only)
        cModel.Log cstrSpacer, cModel.ShowDebug
        cModel.Log PadRight("Exporting relations...", 24), , cModel.ShowDebug
        cModel.Log "", cModel.ShowDebug
        
        intObjCnt = 0
        strObjectPath = strSourcePath & "relations\"
        
        VerifyPath Left(strObjectPath, InStrRev(strObjectPath, "\"))
        ClearOrphanedSourceFiles strObjectPath, dbs.Relations, cModel, "txt"
        
        Dim aRelation As Relation
        For Each aRelation In CurrentDb.Relations
            strName = aRelation.Name
            If Not (strName = "MSysNavPaneGroupsMSysNavPaneGroupToObjects" Or strName = "MSysNavPaneGroupCategoriesMSysNavPaneGroups") Then
                cModel.Log "  " & strName, cModel.ShowDebug
                strName = GetRelationFileName(aRelation)
                modRelation.ExportRelation aRelation, strObjectPath & strName & ".txt"
                intObjCnt = intObjCnt + 1
            End If
        Next aRelation
    
        If cModel.ShowDebug Then
            cModel.Log "[" & intObjCnt & "] relations exported."
        Else
            cModel.Log "[" & intObjCnt & "]"
        End If
    End If
    
    
    ' VBE objects
    If cModel.IncludeVBE Then
        cModel.Log cstrSpacer, cModel.ShowDebug
        cModel.Log PadRight("Exporting VBE...", 24), , cModel.ShowDebug
        cModel.Log "", cModel.ShowDebug
        ExportAllVBE cModel
    End If

    ' Show final output and save log
    cModel.Log cstrSpacer, cModel.ShowDebug
    cModel.Log "Done. (" & Round(Timer - sngStart, 2) & " seconds)"
    cModel.SaveLogFile strSourcePath & "\Export.log"
    
    ' Clean up after completion
    Set m_FSO = Nothing
    
    ' Save version from last
    If GetDBProperty("Last VCS Version") <> GetVCSVersion Then
        SetDBProperty "Last VCS Version", GetVCSVersion
        ' Reload version control so we can run fast save.
        'InitializeVersionControlSystem
    End If
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ExportVBE
' Author    : Adam Waller
' Date      : 5/15/2015
' Purpose   : Exports all objects from the Visual Basic Editor.
'           : (Allows drag and drop to re-import the objects into the IDE)
'---------------------------------------------------------------------------------------
'
Public Sub ExportAllVBE(cModel As IVersionControl)
    
    ' Declare constants locally to avoid need for reference
    'Const vbext_ct_StdModule As Integer = 1
    'Const vbext_ct_MSForm As Integer = 3
    
    Dim cmp As VBIDE.VBComponent
    Dim strExt As String
    Dim strPath As String
    Dim obj_count As Integer
    
    Set colVerifiedPaths = New Collection   ' Reset cache

    
    strPath = cModel.ExportBaseFolder
    VerifyPath strPath
    strPath = strPath & "VBE\"
    
    ' Clear existing files
    ClearTextFilesFromDir strPath, "bas"
    ClearTextFilesFromDir strPath, "frm"
    ClearTextFilesFromDir strPath, "cls"
    
    If VBE.ActiveVBProject.VBComponents.Count > 0 Then
    
        ' Verify path (creating if needed)
        VerifyPath strPath
       
        ' Loop through all components in the active project
        For Each cmp In VBE.ActiveVBProject.VBComponents
            obj_count = obj_count + 1
            strExt = GetVBEExtByType(cmp)
            cmp.Export strPath & cmp.Name & strExt
            cModel.Log "  " & cmp.Name, cModel.ShowDebug
        Next cmp
        
        If cModel.ShowDebug Then
            cModel.Log "[" & obj_count & "] components exported."
        Else
            cModel.Log "[" & obj_count & "]"
        End If
    End If

    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ExportByVBEComponentName
' Author    : Adam Waller
' Date      : 5/15/2015
' Purpose   : Export single object using the VBE component name
'---------------------------------------------------------------------------------------
'
Public Sub ExportByVBEComponent(cmpToExport As VBComponent, cModel As IVersionControl)
    
    Dim intType As AcObjectType
    Dim strFolder As String
    Dim strName As String
    Dim blnSanitize As Boolean
    Dim strFile As String
    
    ' Determine the type of object, and get name of item
    ' in Microsoft Access. (Can be different from VBE)
    With cmpToExport
        Select Case .Type
            Case vbext_ct_StdModule, vbext_ct_ClassModule
                ' Code modules
                intType = acModule
                strName = .Name
                strFolder = "modules\"
            
            Case vbext_ct_Document
                ' Class object (Forms, Reports)
                If Left(.Name, 5) = "Form_" Then
                    intType = acForm
                    strName = Mid(.Name, 6)
                    strFolder = "forms\"
                    blnSanitize = True
                ElseIf Left(.Name, 7) = "Report_" Then
                    intType = acReport
                    strName = Mid(.Name, 8)
                    strFolder = "reports\"
                    blnSanitize = True
                End If
                
        End Select
    End With
    
    DoCmd.Hourglass True
    If intType > 0 Then
        strFolder = cModel.ExportBaseFolder & strFolder
        strFile = strFolder & GetSafeFileName(strName) & ".bas"
        ' Export the single object
        Debug.Print "  Export Path: " & strFile
        ExportObject intType, strName, strFile, cModel
        ' Sanitize object if needed
        If blnSanitize Then SanitizeFile strFile, cModel
    End If
    
    ' Export VBE version
    If cModel.IncludeVBE Then
        strFile = cModel.ExportBaseFolder & "VBE\" & cmpToExport.Name & GetVBEExtByType(cmpToExport)
        If Dir(strFile) <> "" Then Kill strFile
        cmpToExport.Export strFile
    End If
    DoCmd.Hourglass False
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ExportObject
' Author    : Adam Waller
' Date      : 1/23/2019
' Purpose   : Export a database object with optional UCS2-to-UTF-8 conversion.
'---------------------------------------------------------------------------------------
'
Public Sub ExportObject(intType As AcObjectType, strObject As String, strPath As String, cModel As IVersionControl)
        
    Dim blnSkip As Boolean
    Dim strTempFile As String
    Dim strFile As String
    Dim strFolder As String
    Dim dbs As DAO.Database
    
    On Error GoTo ErrHandler
    
    strFolder = Left(strPath, InStrRev(strPath, "\"))
    VerifyPath strFolder
    
    ' Check for fast save
    If cModel.FastSave Then
        Select Case intType
            Case acQuery
                blnSkip = Not (HasMoreRecentChanges(CurrentData.AllQueries(strObject), strPath))
            Case acForm
                blnSkip = Not (HasMoreRecentChanges(CurrentProject.AllForms(strObject), strPath))
            Case acReport
                blnSkip = Not (HasMoreRecentChanges(CurrentProject.AllReports(strObject), strPath))
            Case acMacro
                blnSkip = Not (HasMoreRecentChanges(CurrentProject.AllMacros(strObject), strPath))
        End Select
    End If
    
    If blnSkip Then
        cModel.Log "  (Skipping '" & strObject & "')", cModel.ShowDebug
    Else
        Set dbs = CurrentDb
    
        ' Special options for SQL queries
        If intType = acQuery And cModel.SaveQuerySQL Then
            ' Support for SQL export for queries.
            strFile = strFolder & GetSafeFileName(strObject) & ".sql"
            WriteFile dbs.QueryDefs(strObject).sql, strFile
            cModel.Log "  " & strObject & " (with SQL)", cModel.ShowDebug
            
        ' Log other object
        Else
            cModel.Log "  " & strObject, cModel.ShowDebug
        End If
    
        ' Export object as text (sanitize if needed.)
        Select Case intType
            Case acForm, acReport, acQuery, acMacro
                If CurrentProject.ProjectType = acADP Then
                    ' No UCS conversion needed.
                    Application.SaveAsText intType, strObject, strPath
                Else
                    ' Convert UCS to UTF-8
                    strTempFile = GetTempFile
                    Application.SaveAsText intType, strObject, strTempFile
                    ConvertUcs2Utf8 strTempFile, strPath
                    Kill strTempFile
                End If
                SanitizeFile strPath, cModel
            Case Else
                ' Other object type
                Application.SaveAsText intType, strObject, strPath
        End Select
    End If

    Exit Sub
    
ErrHandler:
    Select Case Err.Number
        Case 2950
            ' Reserved error. Probably couldn't run the SaveAsText command.
            ' (This can happen, for example, when you try to save a data macros on a table that doesn't contain them.)
            Err.Clear
            Resume Next
    Case Else
        ' Unhandled error
        Debug.Print Err.Number & ": " & Err.Description
        Stop
    End Select
    
End Sub


' Import a database object with optional UTF-8-to-UCS2 conversion.
Public Sub ImportObject(obj_type_num As Integer, obj_name As String, file_path As String, _
    Optional Ucs2Convert As Boolean = False)
    
    If Not FSO.FileExists(file_path) Then Exit Sub
    
    If Ucs2Convert Then
        Dim tempFileName As String: tempFileName = modFileAccess.GetTempFile()
        modFileAccess.ConvertUtf8Ucs2 file_path, tempFileName
        Application.LoadFromText obj_type_num, obj_name, tempFileName
        
        FSO.DeleteFile tempFileName
    Else
        Application.LoadFromText obj_type_num, obj_name, file_path
    End If
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ImportObjects
' Author    : Marco Salernitano
' Date      : 14-Nov-2019
' Purpose   : Import multiple objects automatically recognizing their type
'---------------------------------------------------------------------------------------
Public Function ImportObjects(ObjectImportPaths As Collection, Optional simulate As Boolean) As Scripting.Dictionary
    Dim ObjectImportPath    As Variant
    Dim importedCount       As Long
    
    Dim importedMap         As New Scripting.Dictionary
    importedMap.Add "Failed", CLng(0)
    
    For Each ObjectImportPath In ObjectImportPaths
        If FSO.FileExists(ObjectImportPath) Then
            Dim importedObject      As String
            importedObject = RecogAndImportObject((ObjectImportPath), simulate, True)
            If importedObject = "" Then importedObject = "Failed"
            If importedMap.Exists(importedObject) Then
                importedMap(importedObject) = importedMap(importedObject) + 1
            Else
                importedMap.Add importedObject, CLng(1)
            End If
        End If
    Next
    Set ImportObjects = importedMap
End Function

'---------------------------------------------------------------------------------------
' Procedure : CheckMarkerInFile
' Author    : Marco Salernitano
' Date      : 14-Nov-2019
' Purpose   : search for a given marker in the given top lines of a file
'---------------------------------------------------------------------------------------
Public Function CheckMarkerInFile(file_path As String, marker As String, _
    Optional Ucs2Convert As Boolean = False, Optional topLinesToCheck As Long) As Boolean
    
    If Not FSO.FileExists(file_path) Then Exit Function
    Dim tempFileName As String
    If Ucs2Convert Then
        tempFileName = modFileAccess.GetTempFile()
        modFileAccess.ConvertUtf8Ucs2 file_path, tempFileName
    Else
        tempFileName = file_path
    End If
    
    Dim Stream As Scripting.TextStream
    Set Stream = FSO.GetFile(file_path).OpenAsTextStream(ForReading)
    Do While Not Stream.AtEndOfStream
        Dim line As String
        line = Stream.ReadLine
        CheckMarkerInFile = InStr(line, marker) > 0
        If CheckMarkerInFile Then Exit Do
        If topLinesToCheck > 0 Then If Stream.line > topLinesToCheck Then Exit Do
    Loop
    Stream.Close
    If Ucs2Convert Then FSO.DeleteFile tempFileName
End Function

'---------------------------------------------------------------------------------------
' Procedure : RecogAndImportObject
' Author    : Marco Salernitano
' Date      : 14-Nov-2019
' Purpose   : attempt to recognize and import a given file if it is a serialized object
'---------------------------------------------------------------------------------------
 Public Function RecogAndImportObject(ObjectImportPath As String, Optional simulate As Boolean, Optional ShowDebugInfo As Boolean) As String
 
    Dim obj_path_split() As String
    Dim obj_path_part As Variant
    Dim obj_type As Variant
    Dim obj_type_split() As String
    Dim obj_type_label As String
    Dim obj_type_num As Variant
    Dim obj_type_ext As Variant
    Dim obj_type_mark As Variant
    Dim obj_type_topl As Long
    Dim obj_name As String
    Dim ucs2 As Boolean
 
    ' types: type1 , type2 ...
    ' type: label | acObjectType | extensions | markers | toplines
    ' extensions;marker: item1 ; item2; item3 ... (in OR)
    ' be careful to not insert separation characters ,| in markers
    ' sequence order counts
    Const cObjectTypes = _
        "queries|" & acQuery & "|bas;qry;accqry" & "|Operation =;dbMemo|1" & "," & _
        "modules|" & acModule & "|bas;vba;accmod" & "|Option Explicit;Attribute VB|5" & "," & _
        "forms|" & acForm & "|bas;frm;accfrm" & "|Begin Form|10" & "," & _
        "reports|" & acReport & "|bas;rpt;accrpt" & "|Begin Report|10" & "," & _
        "macros|" & acMacro & "|bas;mcr;accmcr" & "|    Action =;    Comment =;    Condition =|5" & "," & _
        "tables|" & acTable & "|txt;acctdt" & "|" & vbTab & "|1" & "," & _
        "tbldefs|" & acModule & "|xml;acctdf" & "|urn:schemas-microsoft-com:officedata|10" & "," & _
        "properties|" & acDatabaseProperties & "|txt;accprp" & "|Connect=|10" & "," & _
        "relations|" & "rel" & "|txt;accrel" & "|Field = Begin|0" & "," & _
        "importspecs|" & "spec" & "|spec;accspc" & "|urn:www.microsoft.com/office/access/imexspec|10" & "," & _
        "references|" & "ref" & "|csv;accref" & "|{"
    
    'search for all types
    For Each obj_type In Split(cObjectTypes, ",")
        obj_type_num = Empty
        obj_type_split = Split(obj_type, "|")
        obj_type_label = obj_type_split(0)
        
'        ' search label in path
'        obj_path_split = Split(ObjectImportPath, "\")
'        For Each obj_path_part In obj_path_split
'            If InStr(obj_path_part, obj_type_label) <> 0 Then
'                obj_type_num = obj_type_split(1)
'                Exit For
'            End If
'        Next
        'if label is not found in path then check for file extension(s)
        If IsEmpty(obj_type_num) Then
            For Each obj_type_ext In Split(obj_type_split(2), ";")
                If obj_type_ext = FSO.GetExtensionName(ObjectImportPath) Then
                    obj_type_num = obj_type_split(1)
                    Exit For
                End If
            Next
        End If
        If Not IsEmpty(obj_type_num) Then
            Select Case obj_type_label
                Case "modules", "macros", "reports", "forms"
                    ucs2 = False
                Case Else
                    ucs2 = modFileAccess.UsingUcs2
            End Select
            ' if found the label in path or the extension then
            ' check for marker as confirmation
            Dim markerFound As Boolean
            markerFound = False
            obj_type_topl = Val(obj_type_split(4))
            For Each obj_type_mark In Split(obj_type_split(3), ";")
                markerFound = CheckMarkerInFile(ObjectImportPath, (obj_type_mark), ucs2, obj_type_topl)
                If markerFound Then Exit For
            Next
            If markerFound Then
                Exit For
            Else
                obj_type_num = Empty
            End If
        End If
    Next
    
    'if a type matched
    If Not IsEmpty(obj_type_num) Then
        If IsNumeric(obj_type_num) Then
            If Not simulate Then ImportObject Val(obj_type_num), FSO.GetBaseName(ObjectImportPath), CStr(ObjectImportPath), ucs2
        Else
            Select Case obj_type_num
                Case "ref"
                    'call ref importer  <<<<<<<<<<<<<<<<<<<
                Case Else ' "rel", "spec"
                    Debug.Print PadRight("No function '" & obj_type_num & "' for:", 24); ObjectImportPath
                    Exit Function
            End Select
        End If
        RecogAndImportObject = obj_type_label
        If ShowDebugInfo Then Debug.Print PadRight("Imported in " & obj_type_label & ":", 24); ObjectImportPath
    Else
        Debug.Print PadRight("Unknown object type in:", 24); ObjectImportPath
    End If
 End Function

' Main entry point for IMPORT. Import all forms, reports, queries,
' macros, modules, and lookup tables from `source` folder under the
' database's folder.
Public Sub ImportAllSource(Optional ShowDebugInfo As Boolean = False, Optional source_path As String)
    
    Dim dbs As DAO.Database
'    Dim source_path As String
    Dim obj_path As String
    Dim obj_type As Variant
    Dim obj_type_split() As String
    Dim obj_type_label As String
    Dim obj_type_num As Integer
    Dim obj_count As Integer
    Dim FileName As String
    Dim obj_name As String
    Dim ucs2 As Boolean

    ' Make sure we are not trying to import into our runing code.
    If CurrentProject.Name = CodeProject.Name Then
        MsgBox "Module " & obj_name & "Code modules cannot be updated while running." & vbCrLf & "Please update manually", vbCritical, "Unable to import source"
        Exit Sub
    End If

    Set dbs = CurrentDb

    If source_path = "" Then source_path = ProjectPath() & "source\"
    If Not FSO.FolderExists(source_path) Then
        MsgBox "No source found at:" & vbCrLf & source_path, vbExclamation, "Import failed"
        Exit Sub
    End If

    Debug.Print
    
    If Not modReference.ImportReferences(source_path) Then
        Debug.Print "Info: no references file in " & source_path
        Debug.Print
    End If

    obj_path = source_path & "queries\"
    FileName = Dir(obj_path & "*.bas")
    Dim tempFilePath As String: tempFilePath = modFileAccess.GetTempFile()
    If Len(FileName) > 0 Then
        Debug.Print PadRight("Importing queries...", 24);
        obj_count = 0
        Do Until Len(FileName) = 0
            DoEvents
            obj_name = Mid(FileName, 1, InStrRev(FileName, ".") - 1)
            ImportObject acQuery, obj_name, obj_path & FileName, modFileAccess.UsingUcs2
            ExportObject acQuery, obj_name, tempFilePath, Nothing
            ImportObject acQuery, obj_name, tempFilePath, modFileAccess.UsingUcs2
            obj_count = obj_count + 1
            FileName = Dir()
        Loop
        Debug.Print "[" & obj_count & "]"
    End If
    
    If FSO.FileExists(tempFilePath) Then Kill tempFilePath

    ' restore table definitions
    obj_path = source_path & "tbldefs\"
    FileName = Dir(obj_path & "*.sql")
    If Len(FileName) > 0 Then
        Debug.Print PadRight("Importing tabledefs...", 24);
        obj_count = 0
        Do Until Len(FileName) = 0
            obj_name = Mid(FileName, 1, InStrRev(FileName, ".") - 1)
            If ShowDebugInfo Then
                If obj_count = 0 Then
                    Debug.Print
                End If
                Debug.Print "  [debug] table " & obj_name;
                Debug.Print
            End If
            'modTable.ImportTableDef CStr(obj_name), obj_path
            obj_count = obj_count + 1
            FileName = Dir()
        Loop
        Debug.Print "[" & obj_count & "]"
    End If
    
    
    ' restore linked tables - we must have access to the remote store to import these!
    FileName = Dir(obj_path & "*.LNKD")
    If Len(FileName) > 0 Then
        Debug.Print PadRight("Importing Linked tabledefs...", 24);
        obj_count = 0
        Do Until Len(FileName) = 0
            obj_name = Mid(FileName, 1, InStrRev(FileName, ".") - 1)
            If ShowDebugInfo Then
                If obj_count = 0 Then
                    Debug.Print
                End If
                Debug.Print "  [debug] table " & obj_name;
                Debug.Print
            End If
            modTable.ImportLinkedTable CStr(obj_name), obj_path
            obj_count = obj_count + 1
            FileName = Dir()
        Loop
        Debug.Print "[" & obj_count & "]"
    End If
    
    
    
    ' NOW we may load data
    obj_path = source_path & "tables\"
    FileName = Dir(obj_path & "*.txt")
    If Len(FileName) > 0 Then
        Debug.Print PadRight("Importing tables...", 24);
        obj_count = 0
        Do Until Len(FileName) = 0
            DoEvents
            obj_name = Mid(FileName, 1, InStrRev(FileName, ".") - 1)
            modTable.ImportTableData CStr(obj_name), obj_path
            obj_count = obj_count + 1
            FileName = Dir()
        Loop
        Debug.Print "[" & obj_count & "]"
    End If
    
    'load Data Macros - not DRY!
    obj_path = source_path & "tbldefs\"
    FileName = Dir(obj_path & "*.xml")
    If Len(FileName) > 0 Then
        Debug.Print PadRight("Importing Data Macros...", 24);
        obj_count = 0
        Do Until Len(FileName) = 0
            DoEvents
            obj_name = Mid(FileName, 1, InStrRev(FileName, ".") - 1)
            'modTable.ImportTableData CStr(obj_name), obj_path
            modDataMacro.ImportDataMacros obj_name, obj_path
            obj_count = obj_count + 1
            FileName = Dir()
        Loop
        Debug.Print "[" & obj_count & "]"
    End If
    

        'import Data Macros
    

    For Each obj_type In Split( _
        "forms|" & acForm & "," & _
        "reports|" & acReport & "," & _
        "macros|" & acMacro & "," & _
        "modules|" & acModule _
        , "," _
    )
        obj_type_split = Split(obj_type, "|")
        obj_type_label = obj_type_split(0)
        obj_type_num = Val(obj_type_split(1))
        obj_path = source_path & obj_type_label & "\"
        
        
    
        FileName = Dir(obj_path & "*.bas")
        If Len(FileName) > 0 Then
            Debug.Print PadRight("Importing " & obj_type_label & "...", 24);
            obj_count = 0
            Do Until Len(FileName) = 0
                ' DoEvents no good idea!
                obj_name = Mid(FileName, 1, InStrRev(FileName, ".") - 1)
                Select Case obj_type_label
                    Case "modules", "macros", "reports", "forms"
                        ucs2 = False
                    Case Else
                        ucs2 = modFileAccess.UsingUcs2
                End Select
                
                ImportObject obj_type_num, obj_name, obj_path & FileName, ucs2
                obj_count = obj_count + 1

                FileName = Dir()
            Loop
            Debug.Print "[" & obj_count & "]"
        
        End If
    Next
    
    'import Print Variables
    Debug.Print PadRight("Importing Print Vars...", 24);
    obj_count = 0
    
    obj_path = source_path & "reports\"
    FileName = Dir(obj_path & "*.pv")
    Do Until Len(FileName) = 0
        DoEvents
        obj_name = Mid(FileName, 1, InStrRev(FileName, ".") - 1)
        modReport.ImportPrintVars obj_name, obj_path & FileName
        obj_count = obj_count + 1
        FileName = Dir()
    Loop
    Debug.Print "[" & obj_count & "]"
    
    'import relations
    Debug.Print PadRight("Importing Relations...", 24);
    obj_count = 0
    obj_path = source_path & "relations\"
    FileName = Dir(obj_path & "*.txt")
    Do Until Len(FileName) = 0
        DoEvents
        modRelation.ImportRelation obj_path & FileName
        obj_count = obj_count + 1
        FileName = Dir()
    Loop
    Debug.Print "[" & obj_count & "]"
    DoEvents
    Debug.Print "Done."
End Sub


' Main entry point for ImportProject.
' Drop all forms, reports, queries, macros, modules.
' execute ImportAllSource.
Public Sub ImportProject(Optional ShowDebugInfo As Boolean = False, Optional source_path As String)
    
    On Error GoTo errorHandler

    ' Make sure we are not trying to delete our runing code.
    If CurrentProject.Name = CodeProject.Name Then
        MsgBox "Code modules cannot be removed while running." & vbCrLf & "Please update manually", vbCritical, "Unable to import source"
        Exit Sub
    End If


    If MsgBox("This action will delete all existing: " & vbCrLf & _
              vbCrLf & _
              Chr(149) & " Tables" & vbCrLf & _
              Chr(149) & " Forms" & vbCrLf & _
              Chr(149) & " Macros" & vbCrLf & _
              Chr(149) & " Modules" & vbCrLf & _
              Chr(149) & " Queries" & vbCrLf & _
              Chr(149) & " Reports" & vbCrLf & _
              vbCrLf & _
              "Are you sure you want to proceed?", vbCritical + vbYesNo, _
              "Import Project") <> vbYes Then
        Exit Sub
    End If

    Dim Db As DAO.Database
    Set Db = CurrentDb
    CloseAllFormsReports

    Debug.Print
    Debug.Print "Deleting Existing Objects"
    Debug.Print
    
    Dim rel As Relation
    For Each rel In CurrentDb.Relations
        If Not (rel.Name = "MSysNavPaneGroupsMSysNavPaneGroupToObjects" Or rel.Name = "MSysNavPaneGroupCategoriesMSysNavPaneGroups") Then
            CurrentDb.Relations.Delete (rel.Name)
        End If
    Next

    Dim dbObject As Object
    For Each dbObject In Db.QueryDefs
        DoEvents
        If Left(dbObject.Name, 1) <> "~" Then
'            Debug.Print dbObject.Name
            Db.QueryDefs.Delete dbObject.Name
        End If
    Next
    
    Dim td As TableDef
    For Each td In CurrentDb.TableDefs
        If Left$(td.Name, 4) <> "MSys" And _
            Left(td.Name, 1) <> "~" Then
            CurrentDb.TableDefs.Delete (td.Name)
        End If
    Next

    Dim objType As Variant
    Dim objTypeArray() As String
    Dim doc As Object
    '
    '  Object Type Constants
    Const OTNAME = 0
    Const OTID = 1

    For Each objType In Split( _
            "Forms|" & acForm & "," & _
            "Reports|" & acReport & "," & _
            "Scripts|" & acMacro & "," & _
            "Modules|" & acModule _
            , "," _
        )
        objTypeArray = Split(objType, "|")
        DoEvents
        For Each doc In Db.Containers(objTypeArray(OTNAME)).Documents
            DoEvents
            If (Left(doc.Name, 1) <> "~") Then
'                Debug.Print doc.Name
                DoCmd.DeleteObject objTypeArray(OTID), doc.Name
            End If
        Next
    Next
    
    Debug.Print "================="
    Debug.Print "Importing Project"
    ImportAllSource ShowDebugInfo, source_path
    GoTo exitHandler

errorHandler:
  Debug.Print "modImportExport.ImportProject: Error #" & Err.Number & vbCrLf & _
               Err.Description

exitHandler:
End Sub

' Main entry point for ImportProject.
' Drop all forms, reports, queries, macros, modules.
' execute ImportAllSource.
Public Function ResetProject(Optional ShowDebugInfo As Boolean = False) As Boolean
    
    On Error GoTo errorHandler

    ' Make sure we are not trying to delete our runing code.
    If CurrentProject.Name = CodeProject.Name Then
        MsgBox "Code modules cannot be removed while running." & vbCrLf & "Please update manually", vbCritical, "Unable to import source"
        Exit Function
    End If


    If MsgBox("This action will delete all existing: " & vbCrLf & _
              vbCrLf & _
              Chr(149) & " Tables" & vbCrLf & _
              Chr(149) & " Forms" & vbCrLf & _
              Chr(149) & " Macros" & vbCrLf & _
              Chr(149) & " Modules" & vbCrLf & _
              Chr(149) & " Queries" & vbCrLf & _
              Chr(149) & " Reports" & vbCrLf & _
              vbCrLf & _
              "Are you sure you want to proceed?", vbCritical + vbYesNo, _
              "Import Project") <> vbYes Then
        Exit Function
    End If

    Dim Db As DAO.Database
    Set Db = CurrentDb
    CloseAllFormsReports

    Debug.Print
    Debug.Print "Deleting Existing Objects"
    Debug.Print
    
    Dim rel As Relation
    For Each rel In CurrentDb.Relations
        If Not (rel.Name = "MSysNavPaneGroupsMSysNavPaneGroupToObjects" Or rel.Name = "MSysNavPaneGroupCategoriesMSysNavPaneGroups") Then
            CurrentDb.Relations.Delete (rel.Name)
        End If
    Next

    Dim dbObject As Object
    For Each dbObject In Db.QueryDefs
        DoEvents
        If Left(dbObject.Name, 1) <> "~" Then
'            Debug.Print dbObject.Name
            Db.QueryDefs.Delete dbObject.Name
        End If
    Next
    
    Dim td As TableDef
    For Each td In CurrentDb.TableDefs
        If Left$(td.Name, 4) <> "MSys" And _
            Left(td.Name, 1) <> "~" Then
            CurrentDb.TableDefs.Delete (td.Name)
        End If
    Next

    Dim objType As Variant
    Dim objTypeArray() As String
    Dim doc As Object
    '
    '  Object Type Constants
    Const OTNAME = 0
    Const OTID = 1

    For Each objType In Split( _
            "Forms|" & acForm & "," & _
            "Reports|" & acReport & "," & _
            "Scripts|" & acMacro & "," & _
            "Modules|" & acModule _
            , "," _
        )
        objTypeArray = Split(objType, "|")
        DoEvents
        For Each doc In Db.Containers(objTypeArray(OTNAME)).Documents
            DoEvents
            If (Left(doc.Name, 1) <> "~") Then
'                Debug.Print doc.Name
                DoCmd.DeleteObject objTypeArray(OTID), doc.Name
            End If
        Next
    Next
    
    Debug.Print "================="
    ResetProject = True
    GoTo exitHandler

errorHandler:
  Debug.Print "modImportExport.ResetProject: Error #" & Err.Number & vbCrLf & _
               Err.Description

exitHandler:
End Function

' Expose for use as function, can be called by a query
Public Function Make()
    ImportProject
End Function


'---------------------------------------------------------------------------------------
' Procedure : FSO
' Author    : Adam Waller
' Date      : 1/18/2019
' Purpose   : Wrapper for file system object.
'---------------------------------------------------------------------------------------
'
Public Function FSO() As Scripting.FileSystemObject
    If m_FSO Is Nothing Then Set m_FSO = New Scripting.FileSystemObject
    Set FSO = m_FSO
End Function