Option Compare Database
Option Explicit
Option Private Module


' List of lookup tables that are part of the program rather than the
' data, to be exported with source code
' Set to "*" to export the contents of all tables
'Only used in ExportAllSource
'Private Const INCLUDE_TABLES = ""
' This is used in ImportAllSource
'Private Const DebugOutput = False
'this is used in ExportAllSource
'Causes the mod code to be exported
'Private Const ArchiveMyself = False

Private Const cstrSpacer As String = "-------------------------------"


' Main entry point for EXPORT. Export all forms, reports, queries,
' macros, modules, and lookup tables to `source` folder under the
' database's folder.
Public Sub ExportAllSource(cModel As IVersionControl)
    
    Dim Db As Object ' DAO.Database
    Dim source_path As String
    Dim obj_path As String
    Dim qry As Object ' DAO.QueryDef
    Dim doc As Object ' DAO.Document
    Dim obj_type As Variant
    Dim obj_type_split() As String
    Dim obj_type_label As String
    Dim obj_type_name As String
    Dim obj_type_num As Integer
    Dim obj_count As Integer
    Dim obj_data_count As Integer
    Dim ucs2 As Boolean
    Dim sngStart As Single
    Dim strName As String
    Dim colADO As New Collection
    Dim varType As Variant
    Dim strData As String
    Dim blnSkipFile As Boolean
    Dim strFile As String
    Dim dteLastCompact As Date
    Dim dteModified As Date

    ' Check Ucs2 status
    ucs2 = modFileAccess.UsingUcs2
    
    ' Option used with fast saves
    If cModel.FastSave Then
        strData = GetDBProperty("InitiatedCompactRepair")
        If IsDate(strData) Then dteLastCompact = CDate(strData)
    End If
    
    Set Db = CurrentDb
    sngStart = Timer
    Set colVerifiedPaths = New Collection   ' Reset cache

    ' Read in options from model
    With cModel
        ShowDebugInfo = .ShowDebug
        source_path = .ExportBaseFolder
    End With

    CloseFormsReports
    'InitUsingUcs2

    modFunctions.VerifyPath source_path

    Debug.Print

    ' Process queries
    
    If CurrentProject.ProjectType = acMDB Then
        ' Standard Access Project
        obj_path = source_path & "queries\"
        modFunctions.ClearTextFilesFromDir obj_path, "bas"
        If ShowDebugInfo Then Debug.Print cstrSpacer
        Debug.Print modFunctions.PadRight("Exporting queries...", 24);
        If ShowDebugInfo Then Debug.Print
        obj_count = 0
        For Each qry In Db.QueryDefs
            DoEvents
            If Left(qry.Name, 1) <> "~" Then
                modFunctions.ExportObject acQuery, qry.Name, obj_path & GetSafeFileName(qry.Name) & ".bas", cModel, ucs2
                obj_count = obj_count + 1
            End If
        Next
        modFunctions.SanitizeTextFiles obj_path, "bas", cModel
        If ShowDebugInfo Then
            Debug.Print "[" & obj_count & "] queries exported."
        Else
            Debug.Print "[" & obj_count & "]"
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
        If DirExists(obj_path & "triggers\") Then
            If Not cModel.FastSave Then modFunctions.ClearTextFilesFromDir obj_path & "triggers\", "sql"
        End If
        
        ' Process triggers
        ExportADPTriggers cModel, source_path & "triggers\"
        
        
        ' Loop through each type, exporting SQL definitions
        For Each varType In colADO
            obj_path = source_path & varType(0) & "\"
            modFunctions.VerifyPath obj_path
            
            
            ''''' Wait to clear tables (or other objects) since we need to check the modified date of the file.
            
            If cModel.FastSave Then
                modFunctions.ClearTextFilesForFastSave obj_path, CStr(varType(1)), CStr(varType(0))
            Else
                modFunctions.ClearTextFilesFromDir obj_path, CStr(varType(1))
            End If
            
            If ShowDebugInfo Then Debug.Print cstrSpacer
            Debug.Print modFunctions.PadRight("Exporting " & varType(0) & "...", 24);
            If ShowDebugInfo Then Debug.Print
            obj_count = 0
            For Each qry In varType(2)
                DoEvents
                blnSkipFile = False
                strFile = obj_path & GetSafeFileName(StripDboPrefix(qry.Name)) & "." & varType(1)
                ' Fast save options
                If cModel.FastSave Then
                    dteModified = GetSQLObjectModifiedDate(qry.Name, varType(0))
                    'dteModified = #1/1/2000#
                    If FileExists(strFile) Then
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
                    If ShowDebugInfo Then Debug.Print "  (Skipping '" & qry.Name & "')"
                Else
                    WriteFile strData, strFile
                    If ShowDebugInfo Then Debug.Print "  " & qry.Name
                End If
                obj_count = obj_count + 1
                ' Check for table/query data export
                If InCollection(cModel.TablesToSaveData, qry.Name) Then
                    DoCmd.OutputTo acOutputServerView, qry.Name, acFormatTXT, obj_path & GetSafeFileName(StripDboPrefix(qry.Name)) & ".txt", False
                    If ShowDebugInfo Then Debug.Print "    Data exported"
                End If
            Next qry
            If ShowDebugInfo Then
                Debug.Print "[" & obj_count & "] " & varType(0) & " exported."
            Else
                Debug.Print "[" & obj_count & "]"
            End If
        Next varType
    End If

    ' Clear the cached variables
    GetSQLObjectModifiedDate "", ""
    
    If CurrentProject.ProjectType = acMDB Then
        ' Access Database
        For Each obj_type In Split( _
            "forms|Forms|" & acForm & "," & _
            "reports|Reports|" & acReport & "," & _
            "macros|Scripts|" & acMacro & "," & _
            "modules|Modules|" & acModule _
            , ",")
            
            obj_type_split = Split(obj_type, "|")
            obj_type_label = obj_type_split(0)
            obj_type_name = obj_type_split(1)
            obj_type_num = Val(obj_type_split(2))
            obj_path = source_path & obj_type_label & "\"
            obj_count = 0
            If cModel.FastSave Then
                modFunctions.ClearTextFilesForFastSave obj_path, "bas", obj_type_label
            Else
                modFunctions.ClearTextFilesFromDir obj_path, "bas"
            End If
            If ShowDebugInfo Then Debug.Print cstrSpacer
            Debug.Print modFunctions.PadRight("Exporting " & obj_type_label & "...", 24);
            If ShowDebugInfo Then Debug.Print
            For Each doc In Db.Containers(obj_type_name).Documents
                DoEvents
                If (Left(doc.Name, 1) <> "~") Then
                    If obj_type_label = "modules" Then
                        ucs2 = False
                    Else
                        ucs2 = modFileAccess.UsingUcs2
                    End If
                    modFunctions.ExportObject obj_type_num, doc.Name, obj_path & doc.Name & ".bas", cModel, ucs2
                    
                    If obj_type_label = "reports" Then
                        If cModel.SavePrintVars Then modReport.ExportPrintVars doc.Name, obj_path & doc.Name & ".pv"
                    End If
                    
                    obj_count = obj_count + 1
                End If
            Next
            If ShowDebugInfo Then
                Debug.Print "[" & obj_count & "] " & obj_type_label & " exported."
            Else
                Debug.Print "[" & obj_count & "]"
            End If
    
            If obj_type_label <> "modules" Then
                modFunctions.SanitizeTextFiles obj_path, "bas", cModel
            End If
        Next
    Else
        ' ADP Project
        
        Set colADO = New Collection
        With colADO
            .Add Array("forms", CurrentProject.AllForms, acForm)
            .Add Array("macros", CurrentProject.AllMacros, acMacro)
            .Add Array("modules", CurrentProject.AllModules, acModule)
            .Add Array("reports", CurrentProject.AllReports, acReport)
        End With
        
        For Each varType In colADO
    
            obj_type_label = varType(0)
            obj_type_num = varType(2)
            obj_path = source_path & obj_type_label & "\"
            obj_count = 0
            If cModel.FastSave Then
                modFunctions.ClearTextFilesForFastSave obj_path, "bas", obj_type_label
            Else
                modFunctions.ClearTextFilesFromDir obj_path, "bas"
            End If
            If ShowDebugInfo Then Debug.Print cstrSpacer
            Debug.Print modFunctions.PadRight("Exporting " & obj_type_label & "...", 24);
            If ShowDebugInfo Then Debug.Print
            For Each doc In varType(1)
                DoEvents
                ' Build file name
                strFile = obj_path & GetSafeFileName(StripDboPrefix(doc.Name)) & ".bas"
                ' Check for fast save option
                blnSkipFile = False
                If (cModel.FastSave And (obj_type_num = acForm Or obj_type_num = acReport)) Then
                    If Not HasMoreRecentChanges(doc, strFile) Then
                        ' Yes, we can really skip the export on this object.
                        blnSkipFile = True
                    End If
                
'                    ' Check modified dates on object
'                    dteCreated = doc.DateCreated
'                    dteModified = doc.DateModified
'                    If DatesClose(dteCreated, dteModified) Then
'                        'and also last compact/repair date (which resets these dates)
'                        If DatesClose(dteCreated, dteLastCompact, 20) Then
'                            ' Check change flag from previous compact/repair
'                            If (GetChangeFlag(doc, 0) = 0) Then
'                                ' Make sure repository file hasn't changed
'                                If Not HasMoreRecentChanges(doc, strFile) Then
'                                    ' Yes, we can really skip the export on this object.
'                                    blnSkipFile = True
'                                End If
'                            End If
'                       End If
'                    End If

                End If
                If blnSkipFile Then
                    If ShowDebugInfo Then Debug.Print "  (Skipping '" & doc.Name & "')"
                Else
                    modFunctions.ExportObject obj_type_num, doc.Name, strFile, cModel, ucs2
                    If obj_type_label <> "modules" Then SanitizeFile obj_path, GetSafeFileName(StripDboPrefix(doc.Name)) & ".bas", "bas", cModel
                    If obj_type_label = "reports" Then
                        If cModel.SavePrintVars Then modReport.ExportPrintVars doc.Name, obj_path & GetSafeFileName(StripDboPrefix(doc.Name)) & ".pv"
                    End If
                    ' Skip this step if we aren't using fast saves.
                    If cModel.FastSave Then
                        ' Reset change flag to unchanged since last export.
                        If GetChangeFlag(doc, 0) > 0 Then SetChangeFlag doc, 0
                    End If
                End If
                obj_count = obj_count + 1
            Next doc
            
            If ShowDebugInfo Then
                Debug.Print "[" & obj_count & "] " & obj_type_label & " exported."
            Else
                Debug.Print "[" & obj_count & "]"
            End If
    
        Next varType
    End If
    
    
    If ShowDebugInfo Then Debug.Print cstrSpacer
    Debug.Print modFunctions.PadRight("Exporting references...", 24);
    If ShowDebugInfo Then Debug.Print
    modReference.ExportReferences source_path
    Debug.Print modFunctions.PadRight("Exporting properties...", 24);
    If ShowDebugInfo Then Debug.Print
    modProperty.ExportProperties source_path
    

'-------------------------mdb table export------------------------
    
    If CurrentProject.ProjectType = acMDB Then
        obj_path = source_path & "tables\"
        modFunctions.ClearTextFilesFromDir obj_path, "txt"
        
        Dim td As TableDef
        Dim tds As TableDefs
        Set tds = Db.TableDefs
    
        If cModel.TablesToSaveData.Count > 0 Then
            ' Only create this folder if we are actually saving table data
            modFunctions.MkDirIfNotExist Left(obj_path, InStrRev(obj_path, "\"))
        End If
        
        obj_type_label = "tables"
        obj_type_name = "Table_Def"
        obj_type_num = acTable
        obj_path = source_path & obj_type_label & "\"
        obj_count = 0
        obj_data_count = 0
        
        ' Clear any existing files
        If cModel.FastSave Then
            modFunctions.ClearTextFilesForFastSave obj_path, "LNKD", obj_type_label
            modFunctions.ClearTextFilesForFastSave obj_path, "sql", obj_type_label
            modFunctions.ClearTextFilesForFastSave obj_path, "xml", obj_type_label
        Else
            modFunctions.ClearTextFilesFromDir obj_path, "LNKD"
            modFunctions.ClearTextFilesFromDir obj_path, "sql"
            modFunctions.ClearTextFilesFromDir obj_path, "xml"
        End If
        
        If ShowDebugInfo Then Debug.Print cstrSpacer
        Debug.Print modFunctions.PadRight("Exporting " & obj_type_label & "...", 24);
        If ShowDebugInfo Then Debug.Print
        
        For Each td In tds
            ' This is not a system table
            ' this is not a temporary table
            If Left$(td.Name, 4) <> "MSys" And _
                Left$(td.Name, 1) <> "~" Then
                modFunctions.VerifyPath Left(obj_path, InStrRev(obj_path, "\"))
                If Len(td.connect) = 0 Then ' this is not an external table
                    modTable.ExportTableDef Db, td, td.Name, obj_path, cModel
                    If InCollection(cModel.TablesToSaveData, "*") Then
                        DoEvents
                        modTable.ExportTableData CStr(td.Name), source_path & "tables\"
                        If Len(Dir(source_path & "tables\" & td.Name & ".txt")) > 0 Then
                            obj_data_count = obj_data_count + 1
                        End If
                    ElseIf InCollection(cModel.TablesToSaveData, td.Name) Then
                        DoEvents
                        On Error GoTo Err_TableNotFound
                        modTable.ExportTableData CStr(td.Name), source_path & "tables\"
                        obj_data_count = obj_data_count + 1
Err_TableNotFound:
                        
                    'else don't export table data
                    End If
                    If ShowDebugInfo Then Debug.Print "  " & td.Name
    
                Else
                    modTable.ExportLinkedTable td.Name, obj_path, cModel
                End If
                
                obj_count = obj_count + 1
                
            End If
        Next
        
        If ShowDebugInfo Then
            Debug.Print "[" & obj_count & "] tbldefs exported."
        Else
            Debug.Print "[" & obj_count & "]"
        End If
    
        ' Export relationships (MDB only)
        If ShowDebugInfo Then Debug.Print cstrSpacer
        Debug.Print modFunctions.PadRight("Exporting Relations...", 24);
        If ShowDebugInfo Then Debug.Print
        
        obj_count = 0
        obj_path = source_path & "relations\"
        
        modFunctions.ClearTextFilesFromDir obj_path, "txt"
        
        Dim aRelation As Relation
        For Each aRelation In CurrentDb.Relations
            strName = aRelation.Name
            If Not (strName = "MSysNavPaneGroupsMSysNavPaneGroupToObjects" Or strName = "MSysNavPaneGroupCategoriesMSysNavPaneGroups") Then
                modFunctions.VerifyPath Left(obj_path, InStrRev(obj_path, "\"))
                ' Replace slashes with division character to support linked table relationships
                'http://stackoverflow.com/questions/10708334/how-can-i-create-files-on-windows-with-embedded-slashes-using-python
                'strName = Replace(strName, "\", ChrW(2215))
                ' On second thought, just use table name.
                strName = GetRelationFileName(aRelation)
                modRelation.ExportRelation aRelation, obj_path & strName & ".txt"
                obj_count = obj_count + 1
            End If
        Next aRelation
    
        If ShowDebugInfo Then
            Debug.Print "[" & obj_count & "] relations exported."
        Else
            Debug.Print "[" & obj_count & "]"
        End If
    End If
    

    ' VBE objects
    If cModel.IncludeVBE Then ExportAllVBE cModel

    If ShowDebugInfo Then Debug.Print cstrSpacer
    Debug.Print "Done. (" & Round(Timer - sngStart, 2) & " seconds)"
    
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
    
    ShowDebugInfo = cModel.ShowDebug
    Set colVerifiedPaths = New Collection   ' Reset cache
    
    If ShowDebugInfo Then Debug.Print cstrSpacer
    Debug.Print modFunctions.PadRight("Exporting VBE...", 24);
    If ShowDebugInfo Then Debug.Print
    
    strPath = cModel.ExportBaseFolder
    modFunctions.VerifyPath strPath
    strPath = strPath & "VBE\"
    
    ' Clear existing files
    modFunctions.ClearTextFilesFromDir strPath, "bas"
    modFunctions.ClearTextFilesFromDir strPath, "frm"
    modFunctions.ClearTextFilesFromDir strPath, "cls"
    
    If VBE.ActiveVBProject.VBComponents.Count > 0 Then
    
        ' Verify path (creating if needed)
        modFunctions.VerifyPath strPath
       
        ' Loop through all components in the active project
        For Each cmp In VBE.ActiveVBProject.VBComponents
            obj_count = obj_count + 1
            strExt = GetVBEExtByType(cmp)
            cmp.Export strPath & cmp.Name & strExt
            If ShowDebugInfo Then Debug.Print "  " & cmp.Name
        Next cmp
        
        If ShowDebugInfo Then
            Debug.Print "[" & obj_count & "] components exported."
        Else
            Debug.Print "[" & obj_count & "]"
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
    
    Dim intType As Integer
    Dim strFolder As String
    Dim strName As String
    Dim blnUcs As Boolean
    Dim blnSanitize As Boolean
    
    ' Determine the type of object, and get name of item
    ' in Microsoft Access. (Can be different from VBE)
    With cmpToExport
        Select Case .Type
            Case vbext_ct_StdModule, vbext_ct_ClassModule
                ' Code modules
                intType = acModule
                strName = GetSafeFileName(.Name)
                strFolder = "modules\"
            
            Case vbext_ct_Document
                ' Class object (Forms, Reports)
                If Left(.Name, 5) = "Form_" Then
                    intType = acForm
                    strName = GetSafeFileName(Mid(.Name, 6))
                    strFolder = "forms\"
                    blnSanitize = True
                ElseIf Left(.Name, 7) = "Report_" Then
                    intType = acReport
                    strName = GetSafeFileName(Mid(.Name, 8))
                    strFolder = "reports\"
                    blnSanitize = True
                End If
                
        End Select
    End With
    
    DoCmd.Hourglass True
    If intType > 0 Then
        strFolder = cModel.ExportBaseFolder & strFolder
        ' Export the single object
        ExportObject intType, strName, strFolder & strName & ".bas", cModel, blnUcs
        ' Sanitize object if needed
        If blnSanitize Then SanitizeFile strFolder, strName & ".bas", "bas", cModel
    End If
    
    ' Export VBE version
    If cModel.IncludeVBE Then
        strName = cModel.ExportBaseFolder & "VBE\" & cmpToExport.Name & GetVBEExtByType(cmpToExport)
        If Dir(strName) <> "" Then Kill strName
        cmpToExport.Export strName
    End If
    DoCmd.Hourglass False
    
End Sub


' Main entry point for IMPORT. Import all forms, reports, queries,
' macros, modules, and lookup tables from `source` folder under the
' database's folder.
Public Sub ImportAllSource(Optional ShowDebugInfo As Boolean = False)
    Dim Db As DAO.Database
    Dim fso As New Scripting.FileSystemObject
    Dim source_path As String
    Dim obj_path As String
    Dim obj_type As Variant
    Dim obj_type_split() As String
    Dim obj_type_label As String
    Dim obj_type_num As Integer
    Dim obj_count As Integer
    Dim fileName As String
    Dim obj_name As String
    Dim ucs2 As Boolean

    ' Make sure we are not trying to import into our runing code.
    If CurrentProject.Name = CodeProject.Name Then
        MsgBox "Module " & obj_name & "Code modules cannot be updated while running." & vbCrLf & "Please update manually", vbCritical, "Unable to import source"
        Exit Sub
    End If

    Set Db = CurrentDb

    CloseFormsReports
    'InitUsingUcs2

    source_path = modFunctions.ProjectPath() & "source\"
    If Not fso.FolderExists(source_path) Then
        MsgBox "No source found at:" & vbCrLf & source_path, vbExclamation, "Import failed"
        Exit Sub
    End If

    Debug.Print
    
    If Not modReference.ImportReferences(source_path) Then
        Debug.Print "Info: no references file in " & source_path
        Debug.Print
    End If

    obj_path = source_path & "queries\"
    fileName = Dir(obj_path & "*.bas")
    Dim tempFilePath As String: tempFilePath = modFileAccess.TempFile()
    If Len(fileName) > 0 Then
        Debug.Print modFunctions.PadRight("Importing queries...", 24);
        obj_count = 0
        Do Until Len(fileName) = 0
            DoEvents
            obj_name = Mid(fileName, 1, InStrRev(fileName, ".") - 1)
            modFunctions.ImportObject acQuery, obj_name, obj_path & fileName, modFileAccess.UsingUcs2
            modFunctions.ExportObject acQuery, obj_name, tempFilePath, Nothing, modFileAccess.UsingUcs2
            modFunctions.ImportObject acQuery, obj_name, tempFilePath, modFileAccess.UsingUcs2
            obj_count = obj_count + 1
            fileName = Dir()
        Loop
        Debug.Print "[" & obj_count & "]"
    End If
    
    modFunctions.DelIfExist tempFilePath

    ' restore table definitions
    obj_path = source_path & "tbldefs\"
    fileName = Dir(obj_path & "*.sql")
    If Len(fileName) > 0 Then
        Debug.Print modFunctions.PadRight("Importing tabledefs...", 24);
        obj_count = 0
        Do Until Len(fileName) = 0
            obj_name = Mid(fileName, 1, InStrRev(fileName, ".") - 1)
            If ShowDebugInfo Then
                If obj_count = 0 Then
                    Debug.Print
                End If
                Debug.Print "  [debug] table " & obj_name;
                Debug.Print
            End If
            modTable.ImportTableDef CStr(obj_name), obj_path
            obj_count = obj_count + 1
            fileName = Dir()
        Loop
        Debug.Print "[" & obj_count & "]"
    End If
    
    
    ' restore linked tables - we must have access to the remote store to import these!
    fileName = Dir(obj_path & "*.LNKD")
    If Len(fileName) > 0 Then
        Debug.Print modFunctions.PadRight("Importing Linked tabledefs...", 24);
        obj_count = 0
        Do Until Len(fileName) = 0
            obj_name = Mid(fileName, 1, InStrRev(fileName, ".") - 1)
            If ShowDebugInfo Then
                If obj_count = 0 Then
                    Debug.Print
                End If
                Debug.Print "  [debug] table " & obj_name;
                Debug.Print
            End If
            modTable.ImportLinkedTable CStr(obj_name), obj_path
            obj_count = obj_count + 1
            fileName = Dir()
        Loop
        Debug.Print "[" & obj_count & "]"
    End If
    
    
    
    ' NOW we may load data
    obj_path = source_path & "tables\"
    fileName = Dir(obj_path & "*.txt")
    If Len(fileName) > 0 Then
        Debug.Print modFunctions.PadRight("Importing tables...", 24);
        obj_count = 0
        Do Until Len(fileName) = 0
            DoEvents
            obj_name = Mid(fileName, 1, InStrRev(fileName, ".") - 1)
            modTable.ImportTableData CStr(obj_name), obj_path
            obj_count = obj_count + 1
            fileName = Dir()
        Loop
        Debug.Print "[" & obj_count & "]"
    End If
    
    'load Data Macros - not DRY!
    obj_path = source_path & "tbldefs\"
    fileName = Dir(obj_path & "*.xml")
    If Len(fileName) > 0 Then
        Debug.Print modFunctions.PadRight("Importing Data Macros...", 24);
        obj_count = 0
        Do Until Len(fileName) = 0
            DoEvents
            obj_name = Mid(fileName, 1, InStrRev(fileName, ".") - 1)
            'modTable.ImportTableData CStr(obj_name), obj_path
            modMacro.ImportDataMacros obj_name, obj_path
            obj_count = obj_count + 1
            fileName = Dir()
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
        
        
    
        fileName = Dir(obj_path & "*.bas")
        If Len(fileName) > 0 Then
            Debug.Print modFunctions.PadRight("Importing " & obj_type_label & "...", 24);
            obj_count = 0
            Do Until Len(fileName) = 0
                ' DoEvents no good idea!
                obj_name = Mid(fileName, 1, InStrRev(fileName, ".") - 1)
                If obj_type_label = "modules" Then
                    ucs2 = False
                Else
                    ucs2 = modFileAccess.UsingUcs2
                End If
                
                modFunctions.ImportObject obj_type_num, obj_name, obj_path & fileName, ucs2
                obj_count = obj_count + 1

                fileName = Dir()
            Loop
            Debug.Print "[" & obj_count & "]"
        
        End If
    Next
    
    'import Print Variables
    Debug.Print modFunctions.PadRight("Importing Print Vars...", 24);
    obj_count = 0
    
    obj_path = source_path & "reports\"
    fileName = Dir(obj_path & "*.pv")
    Do Until Len(fileName) = 0
        DoEvents
        obj_name = Mid(fileName, 1, InStrRev(fileName, ".") - 1)
        modReport.ImportPrintVars obj_name, obj_path & fileName
        obj_count = obj_count + 1
        fileName = Dir()
    Loop
    Debug.Print "[" & obj_count & "]"
    
    'import relations
    Debug.Print modFunctions.PadRight("Importing Relations...", 24);
    obj_count = 0
    obj_path = source_path & "relations\"
    fileName = Dir(obj_path & "*.txt")
    Do Until Len(fileName) = 0
        DoEvents
        modRelation.ImportRelation obj_path & fileName
        obj_count = obj_count + 1
        fileName = Dir()
    Loop
    Debug.Print "[" & obj_count & "]"
    DoEvents
    Debug.Print "Done."
End Sub


' Main entry point for ImportProject.
' Drop all forms, reports, queries, macros, modules.
' execute ImportAllSource.
Public Sub ImportProject()
    
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
    CloseFormsReports

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

errorHandler:
  Debug.Print "modImportExport.ImportProject: Error #" & Err.Number & vbCrLf & _
               Err.Description

exitHandler:
End Sub


' Expose for use as function, can be called by a query
Public Function Make()
    ImportProject
End Function