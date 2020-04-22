Option Compare Database
Option Explicit
Option Private Module


Private Const cstrSpacer As String = "-------------------------------"
Public Const cintPad As Integer = 25

' Keep a persistent reference to file system object after initializing version control.
' This way we don't have to recreate this object dozens of times while using VCS.
Private m_FSO As Scripting.FileSystemObject


'---------------------------------------------------------------------------------------
' Procedure : ExportAllSource
' Author    : Adam Waller
' Date      : 1/18/2019
' Purpose   : Exports all source files for the current project.
'---------------------------------------------------------------------------------------
'
Public Sub ExportAllSource()
    
    
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
    Dim cOptions As clsOptions
    
    ' Load the project options and reset the logs
    Set cOptions = LoadOptions
    ClearLogs

    ' Run any custom sub before export
    If cOptions.RunBeforeExport <> vbNullString Then RunSubInCurrentProject cOptions.RunBeforeExport

    ' Option used with fast saves
    If cOptions.UseFastSave Then
        strData = GetDBProperty("InitiatedCompactRepair")
        If IsDate(strData) Then dteLastCompact = CDate(strData)
    End If
    
    Set dbs = CurrentDb
    sngStart = Timer
    Set colVerifiedPaths = New Collection   ' Reset cache

    With cOptions
        Log cstrSpacer
        Log "Beginning Export of all Source", False
        Log CurrentProject.Name
        Log "VCS Version " & GetVCSVersion
        If .UseFastSave Then Log "Using Fast Save"
        Log Now()
        Log cstrSpacer
    End With
    
    ' Read in options from model
    strSourcePath = cOptions.GetExportFolder

    ' Make sure we have a path for the source files
    VerifyPath strSourcePath

    ' Display header in debug output
    Debug.Print
    Debug.Print cstrSpacer
    Debug.Print "  Exporting All Source"
    Debug.Print cstrSpacer

    ' Process queries
    
    If CurrentProject.ProjectType = acMDB Then
        ' Standard Access Project
        strObjectPath = strSourcePath & "queries\"
        ClearOrphanedSourceFiles strObjectPath, dbs.QueryDefs, cOptions, "bas", "sql"
        Log cstrSpacer, cOptions.ShowDebug
        Log PadRight("Exporting queries...", cintPad), True, cOptions.ShowDebug
        Log "", cOptions.ShowDebug
        intObjCnt = 0
        For Each qry In dbs.QueryDefs
            If Left(qry.Name, 1) <> "~" Then
                strFile = strObjectPath & GetSafeFileName(qry.Name) & ".bas"
                ExportObject acQuery, qry.Name, strFile, cOptions
                intObjCnt = intObjCnt + 1
            End If
        Next
        If cOptions.ShowDebug Then
            Log "[" & intObjCnt & "] queries exported."
        Else
            Log "[" & intObjCnt & "]"
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
            If Not cOptions.UseFastSave Then ClearTextFilesFromDir strObjectPath & "triggers\", "sql"
        End If
        
        ' Process triggers
        Log cstrSpacer, cOptions.ShowDebug
        Log PadRight("Exporting triggers...", cintPad), True, cOptions.ShowDebug
        Log "", cOptions.ShowDebug
        ExportADPTriggers cOptions, strSourcePath & "triggers\"
        
        ' Loop through each type, exporting SQL definitions
        For Each varType In colADO
            strObjectPath = strSourcePath & varType(0) & "\"
            VerifyPath strObjectPath
            
            
            ''''' Wait to clear tables (or other objects) since we need to check the modified date of the file.
            Set objContainer = varType(2)
            ClearOrphanedSourceFiles strObjectPath, objContainer, cOptions, varType(1)
            
            Log cstrSpacer, cOptions.ShowDebug
            Log PadRight("Exporting " & varType(0) & "...", cintPad), , cOptions.ShowDebug
            Log "", cOptions.ShowDebug
            intObjCnt = 0
            For Each qry In varType(2)
                blnSkipFile = False
                strFile = strObjectPath & GetSafeFileName(StripDboPrefix(qry.Name)) & "." & varType(1)
                ' Fast save options
                If cOptions.UseFastSave Then
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
                    Log "  (Skipping '" & qry.Name & "')", cOptions.ShowDebug
                Else
                    WriteFile strData, strFile
                    Log "  " & qry.Name, cOptions.ShowDebug
                End If
                intObjCnt = intObjCnt + 1
                ' Check for table/query data export
                If cOptions.TablesToExportData.Exists(qry.Name) Then
                    DoCmd.OutputTo acOutputServerView, qry.Name, acFormatTXT, strObjectPath & GetSafeFileName(StripDboPrefix(qry.Name)) & ".txt", False
                    Log "    Data exported", cOptions.ShowDebug
                End If
            Next qry
            If cOptions.ShowDebug Then
                Log "[" & intObjCnt & "] " & varType(0) & " exported."
            Else
                Log "[" & intObjCnt & "]"
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
        ClearOrphanedSourceFiles strObjectPath, objContainer, cOptions, "bas", "pv"
        
        ' Show progress
        Log cstrSpacer, cOptions.ShowDebug
        Log PadRight("Exporting " & strLabel & "...", cintPad), , cOptions.ShowDebug
        Log "", cOptions.ShowDebug
        
        ' Loop through objects in container
        For Each doc In objContainer
            If (Left(doc.Name, 1) <> "~") Then
                ' Get file name (without extension)
                strFile = strObjectPath & StripDboPrefix(GetSafeFileName(doc.Name))
                ExportObject intType, doc.Name, strFile & ".bas", cOptions
                If intType = acReport Then
                    If cOptions.SavePrintVars Then ExportPrintVars doc.Name, strFile & ".pv", cOptions
                End If
                intObjCnt = intObjCnt + 1
            End If
        Next
        
        ' Show total number of objects
        If cOptions.ShowDebug Then
            Log "[" & intObjCnt & "] " & strLabel & " exported."
        Else
            Log "[" & intObjCnt & "]"
        End If

    Next varType

    ' Export references
    Log cstrSpacer, cOptions.ShowDebug
    Log PadRight("Exporting references...", cintPad), , cOptions.ShowDebug
    Log "", cOptions.ShowDebug
    ExportReferences strSourcePath, cOptions
    
    ' Export database properties
    Log cstrSpacer, cOptions.ShowDebug
    Log PadRight("Exporting properties...", cintPad), , cOptions.ShowDebug
    Log "", cOptions.ShowDebug
    ExportProperties strSourcePath, cOptions
    
    ' Export Import/Export Specifications
    Log cstrSpacer, cOptions.ShowDebug
    Log PadRight("Exporting specs...", cintPad), , cOptions.ShowDebug
    Log "", cOptions.ShowDebug
    ExportSpecs strSourcePath, cOptions
    
    

'-------------------------mdb table export------------------------
    
    If CurrentProject.ProjectType = acMDB Then
                
        Dim td As TableDef
        Dim tds As TableDefs
        Set tds = dbs.TableDefs
    
        If cOptions.TablesToExportData.Count = 0 Then
            strObjectPath = strSourcePath & "tables"
            If FSO.FolderExists(strObjectPath) Then ClearOrphanedSourceFiles strObjectPath & "\", Nothing, cOptions, "txt"
        Else
            ' Only create this folder if we are actually saving table data
            MkDirIfNotExist strSourcePath & "tables\"
            ClearOrphanedSourceFiles strSourcePath & "tables\", dbs.TableDefs, cOptions, "txt"
        End If
        
        strLabel = "tbldef"
        strType = "Table_Def"
        intType = acTable
        strObjectPath = strSourcePath & "tbldefs\"
        intObjCnt = 0
        intObjDataCnt = 0
        
        
        ' Verify path and clear any existing files
        VerifyPath Left(strObjectPath, InStrRev(strObjectPath, "\"))
        ClearOrphanedSourceFiles strObjectPath, tds, cOptions, "LNKD", "sql", "xml", "bas"

        Log cstrSpacer, cOptions.ShowDebug
        Log PadRight("Exporting " & strLabel & "...", cintPad), , cOptions.ShowDebug
        Log "", cOptions.ShowDebug
        
        For Each td In tds
            ' This is not a system table
            ' this is not a temporary table
            If Left$(td.Name, 4) <> "MSys" And _
                Left$(td.Name, 1) <> "~" Then
                If Len(td.connect) = 0 Then ' this is not an external table
                    ExportTableDef td.Name, strObjectPath, cOptions
                    If cOptions.TablesToExportData.Exists("*") Then
                        ExportTableData CStr(td.Name), strSourcePath & "tables\", cOptions
                        If Len(Dir(strSourcePath & "tables\" & td.Name & ".txt")) > 0 Then
                            intObjDataCnt = intObjDataCnt + 1
                        End If
                    ElseIf cOptions.TablesToExportData.Exists(td.Name) Then
                        modTable.ExportTableData CStr(td.Name), strSourcePath & "tables\", cOptions
                        intObjDataCnt = intObjDataCnt + 1
                    'else don't export table data
                    End If
    
                Else
                    modTable.ExportLinkedTable td.Name, strObjectPath, cOptions
                End If
                
                intObjCnt = intObjCnt + 1
                
            End If
        Next
        
        If cOptions.ShowDebug Then
            Log "[" & intObjCnt & "] tbldefs exported."
        Else
            Log "[" & intObjCnt & "]"
        End If
    
        ' Export relationships (MDB only)
        Log cstrSpacer, cOptions.ShowDebug
        Log PadRight("Exporting relations...", cintPad), , cOptions.ShowDebug
        Log "", cOptions.ShowDebug
        
        intObjCnt = 0
        strObjectPath = strSourcePath & "relations\"
        
        VerifyPath Left(strObjectPath, InStrRev(strObjectPath, "\"))
        ClearOrphanedSourceFiles strObjectPath, dbs.Relations, cOptions, "txt"
        
        Dim aRelation As Relation
        For Each aRelation In CurrentDb.Relations
            strName = aRelation.Name
            If Not (strName = "MSysNavPaneGroupsMSysNavPaneGroupToObjects" Or strName = "MSysNavPaneGroupCategoriesMSysNavPaneGroups") Then
                Log "  " & strName, cOptions.ShowDebug
                strName = GetRelationFileName(aRelation)
                modRelation.ExportRelation aRelation, strObjectPath & strName & ".txt"
                intObjCnt = intObjCnt + 1
            End If
        Next aRelation
    
        If cOptions.ShowDebug Then
            Log "[" & intObjCnt & "] relations exported."
        Else
            Log "[" & intObjCnt & "]"
        End If
    End If
    
    
    ' VBE objects
    If cOptions.IncludeVBE Then
        Log cstrSpacer, cOptions.ShowDebug
        Log PadRight("Exporting VBE...", cintPad), , cOptions.ShowDebug
        Log "", cOptions.ShowDebug
        ExportAllVBE cOptions
    End If

    ' Show final output and save log
    Log cstrSpacer
    Log "Done. (" & Round(Timer - sngStart, 2) & " seconds)"
    SaveLogFile strSourcePath & "\Export.log"
    
    ' Clean up after completion
    Set m_FSO = Nothing
    
    ' Save version from last
    If GetDBProperty("Last VCS Version") <> GetVCSVersion Then
        SetDBProperty "Last VCS Version", GetVCSVersion
        ' Reload version control so we can run fast save.
        'InitializeVersionControlSystem
    End If

    ' Run any custom sub before export
    If cOptions.RunAfterExport <> vbNullString Then RunSubInCurrentProject cOptions.RunAfterExport

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ExportVBE
' Author    : Adam Waller
' Date      : 5/15/2015
' Purpose   : Exports all objects from the Visual Basic Editor.
'           : (Allows drag and drop to re-import the objects into the IDE)
'---------------------------------------------------------------------------------------
'
Public Sub ExportAllVBE(cOptions As clsOptions)
    
    ' Declare constants locally to avoid need for reference
    'Const vbext_ct_StdModule As Integer = 1
    'Const vbext_ct_MSForm As Integer = 3
    
    Dim cmp As VBIDE.VBComponent
    Dim strExt As String
    Dim strPath As String
    Dim obj_count As Integer
    
    Set colVerifiedPaths = New Collection   ' Reset cache

    
    strPath = cOptions.GetExportFolder
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
            Log "  " & cmp.Name, cOptions.ShowDebug
        Next cmp
        
        If cOptions.ShowDebug Then
            Log "[" & obj_count & "] components exported."
        Else
            Log "[" & obj_count & "]"
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
Public Sub ExportByVBEComponent(cmpToExport As VBComponent, cOptions As clsOptions)
    
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
        strFolder = cOptions.GetExportFolder & strFolder
        strFile = strFolder & GetSafeFileName(strName) & ".bas"
        ' Export the single object
        ExportObject intType, strName, strFile, cOptions
        ' Sanitize object if needed
        If blnSanitize Then SanitizeFile strFile, cOptions
    End If
    
    ' Export VBE version
    If cOptions.IncludeVBE Then
        strFile = cOptions.GetExportFolder & "VBE\" & cmpToExport.Name & GetVBEExtByType(cmpToExport)
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
Public Sub ExportObject(intType As AcObjectType, strObject As String, strPath As String, cOptions As clsOptions)
        
    Dim blnSkip As Boolean
    Dim strTempFile As String
    Dim strFile As String
    Dim strFolder As String
    Dim dbs As DAO.Database
    
    On Error GoTo ErrHandler
    
    strFolder = Left(strPath, InStrRev(strPath, "\"))
    VerifyPath strFolder
    
    ' Check for fast save
    If cOptions.UseFastSave Then
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
        Log "  (Skipping '" & strObject & "')", cOptions.ShowDebug
    Else
        Set dbs = CurrentDb
    
        ' Special options for SQL queries
        If intType = acQuery And cOptions.SaveQuerySQL Then
            ' Support for SQL export for queries.
            strFile = strFolder & GetSafeFileName(strObject) & ".sql"
            WriteFile dbs.QueryDefs(strObject).sql, strFile
            Log "  " & strObject & " (with SQL)", cOptions.ShowDebug
            
        ' Log other object
        Else
            Log "  " & strObject, cOptions.ShowDebug
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
                SanitizeFile strPath, cOptions
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


' Main entry point for IMPORT. Import all forms, reports, queries,
' macros, modules, and lookup tables from `source` folder under the
' database's folder.
Public Sub ImportAllSource(Optional ShowDebugInfo As Boolean = False)
    
    Dim dbs As DAO.Database
    Dim source_path As String
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

    source_path = ProjectPath() & "source\"
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
        Debug.Print PadRight("Importing queries...", cintPad);
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
        Debug.Print PadRight("Importing tabledefs...", cintPad);
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
        Debug.Print PadRight("Importing Linked tabledefs...", cintPad);
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
        Debug.Print PadRight("Importing tables...", cintPad);
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
        Debug.Print PadRight("Importing Data Macros...", cintPad);
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
            Debug.Print PadRight("Importing " & obj_type_label & "...", cintPad);
            obj_count = 0
            Do Until Len(FileName) = 0
                ' DoEvents no good idea!
                obj_name = Mid(FileName, 1, InStrRev(FileName, ".") - 1)
                If obj_type_label = "modules" Then
                    ucs2 = False
                Else
                    ucs2 = modFileAccess.UsingUcs2
                End If
                
                ImportObject obj_type_num, obj_name, obj_path & FileName, ucs2
                obj_count = obj_count + 1

                FileName = Dir()
            Loop
            Debug.Print "[" & obj_count & "]"
        
        End If
    Next
    
    'import Print Variables
    Debug.Print PadRight("Importing Print Vars...", cintPad);
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
    Debug.Print PadRight("Importing Relations...", cintPad);
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
Public Sub ImportProject()
    
    On Error GoTo ErrorHandler

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
    ImportAllSource
    GoTo exitHandler

ErrorHandler:
  Debug.Print "modImportExport.ImportProject: Error #" & Err.Number & vbCrLf & _
               Err.Description

exitHandler:
End Sub


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