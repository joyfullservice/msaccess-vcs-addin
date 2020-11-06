Option Explicit
Option Compare Database
Option Private Module

Public Const JSON_WHITESPACE As Integer = 2
Public Const UTF8_BOM As String = "ï»¿"
Public Const UCS2_BOM As String = "ÿþ"

' Formats used when exporting table data.
Public Enum eTableDataExportFormat
    etdNoData = 0
    etdTabDelimited = 1
    etdXML = 2
    [_Last] = 2
End Enum

' Object types used when determining SQL modification date.
Public Enum eSqlObjectType
    estView
    estStoredProcedure
    estTable
    estTrigger
    estOther
End Enum

' Types of objects that can be exported/imported from a database.
' (Use corresponding constants wherever possible)
' Be careful not to create collisions with two members sharing the
' same value.
Public Enum eDatabaseComponentType
    ' Standard database objects
    edbForm
    edbMacro
    edbModule
    edbQuery
    edbReport
    edbTableDef
    edbTableDataMacro
    edbLinkedTable
    ' ADP specific
    edbAdpTable
    edbAdpFunction
    edbAdpServerView
    edbAdpStoredProcedure
    edbAdpTrigger
    ' Custom object types we are also handling.
    edbTableData
    edbRelation
    edbDbsProperty
    edbProjectProperty
    edbFileProperty
    edbSharedImage
    edbDocument
    edbSavedSpec
    edbImexSpec
    edbNavPaneGroup
    edbTheme
    edbVbeForm
    edbVbeProject
    edbVbeReference
End Enum

' API function to pause processing
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)

' Logging and options classes
Private m_Perf As clsPerformance
Private m_Log As clsLog
Private m_Options As clsOptions

' Keep a persistent reference to file system object after initializing version control.
' This way we don't have to recreate this object dozens of times while using VCS.
Private m_FSO As FileSystemObject


'---------------------------------------------------------------------------------------
' Procedure : SanitizeFile
' Author    : Adam Waller
' Date      : 11/4/2020
' Purpose   : Rewritten version of sanitize function
'---------------------------------------------------------------------------------------
'
Public Sub SanitizeFile(strPath As String)

    Dim strFile As String
    Dim varLines As Variant
    Dim lngLine As Long
    Dim cData As clsConcat
    Dim strLine As String
    Dim strTLine As String
    Dim blnInsideIgnoredBlock As Boolean
    Dim intIndent As Integer
    Dim blnIsReport As Boolean
    Dim sngStartTime As Single

    ' Read text from file, and split into lines
    If HasUcs2Bom(strPath) Then
        strFile = ReadFile(strPath, "Unicode")
    Else
        strFile = ReadFile(strPath)
    End If
    Perf.OperationStart "Sanitize File"
    varLines = Split(strFile, vbCrLf)
    
    ' Delete original file now so we can write it immediately
    ' when the new data has been constructed.
    DeleteFile strPath

    ' Initialize concatenation class to include line breaks
    ' after each line that we add when building new file text.
    sngStartTime = Timer
    Set cData = New clsConcat
    cData.AppendOnAdd = vbCrLf

    ' Using a do loop since we may adjust the line counter
    ' during a loop iteration.
    Do While lngLine <= UBound(varLines)
        
        ' Get unmodified and trimmed line
        strLine = varLines(lngLine)
        strTLine = Trim$(strLine)
        
        ' Improve performance by reducing comparisons
        If Len(strTLine) > 3 And blnInsideIgnoredBlock Then
            ' Ignore this line
        ElseIf Len(strTLine) > 60 And StartsWith(strTLine, "0x") Then
            ' Add binary data line. No need to test this line further.
            cData.Add strLine
        Else
            ' Run the rest of the tests
            Select Case strTLine
            
                ' File version
                Case "Version =21"
                    ' Change version down to 20 to allow import into Access 2010.
                    ' (Haven't seen any significant issues with this.)
                    cData.Add "Version =20"
                
                ' Print settings blocks to ignore
                Case "PrtMip = Begin", _
                    "PrtDevMode = Begin", _
                    "PrtDevModeW = Begin", _
                    "PrtDevNames = Begin", _
                    "PrtDevNamesW = Begin"
                    ' Set flag to ignore lines inside this block.
                    blnInsideIgnoredBlock = True
        
                ' Aggressive sanitize blocks
                Case "GUID = Begin", _
                    "NameMap = Begin", _
                    "dbLongBinary ""DOL"" = Begin", _
                    "dbBinary ""GUID"" = Begin"
                    If Options.AggressiveSanitize Then blnInsideIgnoredBlock = True
                    
                ' Single lines to ignore
                Case "NoSaveCTIWhenDisabled =1"
        
                ' Publish option (used in Queries)
                Case "dbByte ""PublishToWeb"" =""1""", _
                    "PublishOption =1"
                    If Not Options.StripPublishOption Then cData.Add strLine
                
                ' End of block section
                Case "End"
                    If blnInsideIgnoredBlock Then
                        ' Reached the end of the ignored block.
                        blnInsideIgnoredBlock = False
                    Else
                        ' End of included block
                        cData.Add strLine
                    End If
                
                ' See if this file is from a report object
                Case "Begin Report"
                    ' Turn flag on to ignore Right and Bottom lines
                    blnIsReport = True
                    cData.Add strLine
                    
                Case Else
                    If blnInsideIgnoredBlock Then
                        ' Skip if we are in an ignored block
                    ElseIf StartsWith(strTLine, "Checksum =") Then
                        ' Ignore Checksum lines, since they will change.
                    ElseIf StartsWith(strTLine, "BaseInfo =") Then
                        ' BaseInfo is used with combo boxes, similar to RowSource.
                        ' Since the value could span multiple lines, we need to
                        ' check the indent level of the following lines to see how
                        ' many lines to skip.
                        intIndent = GetIndent(strLine)
                        ' Preview the next line, and check the indent level
                        Do While GetIndent(varLines(lngLine + 1)) > intIndent
                            ' Move
                            lngLine = lngLine + 1
                        Loop
                    ElseIf blnIsReport And StartsWith(strLine, "    Right =") Then
                        ' Ignore this line. (Not important, and frequently changes.)
                    ElseIf blnIsReport And StartsWith(strLine, "    Bottom =") Then
                        ' Turn flag back off now that we have ignored these two lines.
                        blnIsReport = False
                    Else
                        ' All other lines will be added.
                        cData.Add strLine
                    End If
            
            End Select
        End If
    
        ' Increment counter to next line
        lngLine = lngLine + 1
    Loop
    
    ' Remove last vbcrlf
    cData.Remove Len(vbCrLf)

    ' Log performance
    Perf.OperationEnd
    Log.Add "    Sanitized in " & Format$(Timer - sngStartTime, "0.00") & " seconds.", Options.ShowDebug
    
    ' Replace original file with sanitized version
    WriteFile cData.GetStr, strPath
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : StartsWith
' Author    : Adam Waller
' Date      : 11/5/2020
' Purpose   : See if a string begins with a specified string.
'---------------------------------------------------------------------------------------
'
Public Function StartsWith(strText As String, strStartsWith As String, Optional Compare As VbCompareMethod = vbBinaryCompare) As Boolean
    StartsWith = (InStr(1, strText, strStartsWith, Compare) = 1)
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetIndent
' Author    : Adam Waller
' Date      : 11/5/2020
' Purpose   : Returns the number of spaces until the first non-space character.
'---------------------------------------------------------------------------------------
'
Public Function GetIndent(strLine As Variant) As Integer
    Dim strChar As String
    strChar = Left$(Trim(strLine), 1)
    If strLine <> vbNullString Then GetIndent = InStr(1, strLine, strChar) - 1
End Function


'---------------------------------------------------------------------------------------
' Procedure : SanitizeXML
' Author    : Adam Waller
' Date      : 4/27/2020
' Purpose   : Remove non-essential data that changes every time the file is exported.
'---------------------------------------------------------------------------------------
'
Public Sub SanitizeXML(strPath As String, Options As clsOptions)

    Dim sngOverall As Single
    Dim sngTimer As Single
    Dim cData As clsConcat
    Dim strText As String
    Dim rxLine As VBScript_RegExp_55.RegExp
    Dim objMatches As VBScript_RegExp_55.MatchCollection
    Dim stmInFile As ADODB.Stream
    Dim blnFound As Boolean
    
    On Error GoTo 0
    
    Set cData = New clsConcat
    Set rxLine = New VBScript_RegExp_55.RegExp
    
    ' Timers to monitor performance
    sngTimer = Timer
    sngOverall = sngTimer
    
    ' Set line search pattern (To remove generated timestamp)
    '<dataroot xmlns:od="urn:schemas-microsoft-com:officedata" generated="2020-04-27T10:28:32">
    rxLine.Pattern = "^\s*(?:<dataroot xmlns:(.+))( generated="".+"")"
    
    ' Open file to read contents line by line.
    Set stmInFile = New ADODB.Stream
    stmInFile.Charset = "UTF-8"
    stmInFile.Open
    stmInFile.LoadFromFile strPath
    strText = stmInFile.ReadText(adReadLine)
    
    
    ' Loop through all the lines in the file
    Do Until stmInFile.EOS
        
        ' Read line from file
        strText = stmInFile.ReadText(adReadLine)
        If Left$(strText, 3) = UTF8_BOM Then strText = Mid$(strText, 4)
        ' Just looking for the first match.
        If Not blnFound Then
        
            ' Check for matching pattern
            If rxLine.Test(strText) Then
                
                ' Return actual matches
                Set objMatches = rxLine.Execute(strText)
                
                ' Replace with empty string
                strText = Replace(strText, objMatches(0).SubMatches(1), vbNullString, , 1)
                blnFound = True
            End If
        End If
        
        ' Add to return string
        cData.Add strText
        cData.Add vbCrLf
    Loop
    
    ' Close and delete original file
    stmInFile.Close
    FSO.DeleteFile strPath
    
    ' Write file all at once, rather than line by line.
    ' (Otherwise the code can bog down with tens of thousands of write operations)
    WriteFile cData.GetStr, strPath

    ' Show stats if debug turned on.
    Log.Add "    Sanitized in " & Format$(Timer - sngOverall, "0.00") & " seconds.", Options.ShowDebug

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ProjectPath
' Author    : Adam Waller
' Date      : 1/25/2019
' Purpose   : Path/Directory of the current database file.
'---------------------------------------------------------------------------------------
'
Public Function ProjectPath() As String
    ProjectPath = CurrentProject.Path
    If Right$(ProjectPath, 1) <> "\" Then ProjectPath = ProjectPath & "\"
End Function


'---------------------------------------------------------------------------------------
' Procedure : MkDirIfNotExist
' Author    : Adam Waller
' Date      : 1/25/2019
' Purpose   : Create folder `Path`. Silently do nothing if it already exists.
'---------------------------------------------------------------------------------------
'
Public Sub MkDirIfNotExist(strPath As String)
    If Not FSO.FolderExists(StripSlash(strPath)) Then FSO.CreateFolder StripSlash(strPath)
End Sub


'---------------------------------------------------------------------------------------
' Procedure : clearfilesbyextension
' Author    : Adam Waller
' Date      : 1/25/2019
' Purpose   : Erase all *.`ext` files in `Path`.
'---------------------------------------------------------------------------------------
'
Public Sub ClearFilesByExtension(ByVal strFolder As String, strExt As String)

    Dim oFile As Scripting.File
    Dim strFolderNoSlash As String
    
    ' While the Dir() function would be simpler, it does not support Unicode.
    strFolderNoSlash = StripSlash(strFolder)
    If FSO.FolderExists(strFolderNoSlash) Then
        For Each oFile In FSO.GetFolder(strFolderNoSlash).Files
            If StrComp(FSO.GetExtensionName(oFile.Name), strExt, vbTextCompare) = 0 Then
                ' Found at least one matching file. Use the wildcard delete.
                FSO.DeleteFile strFolderNoSlash & "\*." & strExt
                Exit Sub
            End If
        Next
    End If
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ClearOrphanedSourceFolders
' Author    : Casper Englund
' Date      : 2020-06-04
' Purpose   : Clears existing source folders that don't have a matching object in the
'           : database.
'---------------------------------------------------------------------------------------
'
Public Sub ClearOrphanedSourceFolders(cType As IDbComponent)
    
    Dim colNames As Collection
    Dim cItem As IDbComponent
    Dim oFolder As Folder
    Dim oSubFolder As Folder
    Dim strSubFolderName As String
    
    ' No orphaned files if the folder doesn't exist.
    If Not FSO.FolderExists(cType.BaseFolder) Then Exit Sub
    
    ' Cache a list of source file names for actual database objects
    Set colNames = New Collection
    For Each cItem In cType.GetAllFromDB
        colNames.Add FSO.GetFileName(cItem.SourceFile)
    Next cItem
    
    Set oFolder = FSO.GetFolder(cType.BaseFolder)
    For Each oSubFolder In oFolder.SubFolders
            
        strSubFolderName = oSubFolder.Name
        ' Remove any subfolder that doesn't have a matching name.
        If Not InCollection(colNames, strSubFolderName) Then
            ' Object not found in database. Remove subfolder.
            oSubFolder.Delete True
            Log.Add "  Removing orphaned folder: " & strSubFolderName, Options.ShowDebug
        End If
        
    Next oSubFolder
    
    ' Remove base folder if we don't have any subfolders in it
    If oFolder.SubFolders.Count = 0 Then oFolder.Delete
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ClearOrphanedSourceFiles
' Author    : Adam Waller
' Date      : 12/14/2016
' Purpose   : Clears existing source files that don't have a matching object in the
'           : database.
'---------------------------------------------------------------------------------------
'
Public Sub ClearOrphanedSourceFiles(cType As IDbComponent, ParamArray StrExtensions())
    
    Dim oFolder As Folder
    Dim oFile As File
    Dim colNames As Collection
    Dim strFile As String
    Dim varExt As Variant
    Dim strPrimaryExt As String
    Dim cItem As IDbComponent
    
    ' No orphaned files if the folder doesn't exist.
    If Not FSO.FolderExists(cType.BaseFolder) Then Exit Sub
    
    ' Cache a list of source file names for actual database objects
    Set colNames = New Collection
    For Each cItem In cType.GetAllFromDB
        colNames.Add FSO.GetFileName(cItem.SourceFile)
    Next cItem
    If colNames.Count > 0 Then strPrimaryExt = "." & FSO.GetExtensionName(colNames(1))
    
    ' Loop through files in folder
    Set oFolder = FSO.GetFolder(cType.BaseFolder)
    For Each oFile In oFolder.Files
    
        ' Check against list of extensions
        For Each varExt In StrExtensions
        
            ' Check for matching extension on wanted list.
            If FSO.GetExtensionName(oFile.Path) = varExt Then
                
                ' Build a file name using the primary extension to
                ' match the list of source files.
                strFile = FSO.GetBaseName(oFile.Name) & strPrimaryExt
                ' Remove any file that doesn't have a matching name.
                If Not InCollection(colNames, strFile) Then
                    ' Object not found in database. Remove file.
                    FSO.DeleteFile oFile.ParentFolder.Path & "\" & oFile.Name, True
                    Log.Add "  Removing orphaned file: " & strFile, Options.ShowDebug
                End If
                
                ' No need to check other extensions since we
                ' already had a match and processed the file.
                Exit For
            End If
        Next varExt
    Next oFile
    
    ' Remove base folder if we don't have any files in it
    If oFolder.Files.Count = 0 Then oFolder.Delete True
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : StripSlash
' Author    : Adam Waller
' Date      : 1/25/2019
' Purpose   : Strip the trailing slash
'---------------------------------------------------------------------------------------
'
Public Function StripSlash(strText As String) As String
    If Right$(strText, 1) = "\" Then
        StripSlash = Left$(strText, Len(strText) - 1)
    Else
        StripSlash = strText
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : InCollection
' Author    : Adam Waller
' Date      : 6/2/2015
' Purpose   : Returns true if the item value is found in the collection
'---------------------------------------------------------------------------------------
'
Public Function InCollection(ByVal MyCol As Collection, ByVal MyValue As Variant) As Boolean
    Dim intCnt As Integer
    For intCnt = 1 To MyCol.Count
        If MyCol(intCnt) = MyValue Then
            InCollection = True
            Exit For
        End If
    Next intCnt
End Function


'---------------------------------------------------------------------------------------
' Procedure : MergeCollection
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Adds a collection into another collection.
'---------------------------------------------------------------------------------------
'
Public Sub MergeCollection(ByRef colOriginal As Collection, ByVal colToAdd As Collection)
    Dim varItem As Variant
    For Each varItem In colToAdd
        colOriginal.Add varItem
    Next varItem
End Sub


'---------------------------------------------------------------------------------------
' Procedure : VerifyPath
' Author    : Adam Waller
' Date      : 8/3/2020
' Purpose   : Verifies that the folder path to a folder or file exists.
'           : Use this to verify the folder path before attempting to write a file.
'---------------------------------------------------------------------------------------
'
Public Sub VerifyPath(strPath As String)
    
    Dim strFolder As String
    Dim varParts As Variant
    Dim intPart As Integer
    Dim strVerified As String
    
    ' Determine if the path is a file or folder
    If Right$(strPath, 1) = "\" Then
        ' Folder name. (Folder names can contain periods)
        strFolder = Left$(strPath, Len(strPath) - 1)
    Else
        ' File name
        strFolder = FSO.GetParentFolderName(strPath)
    End If
    
    ' Check if full path exists.
    If Not FSO.FolderExists(strFolder) Then
        ' Start from the root, and build out full path, creating folders as needed.
        varParts = Split(strFolder, "\")
        ' Make sure the root folder exists. If it doesn't we probably have some other
        ' issue.
        If Not FSO.FolderExists(varParts(0)) Then
            MsgBox2 "Path Not Found", "Could not find the path '" & varParts(0) & "' on this system.", _
                "I was simply trying to verify this path: " & strFolder, vbExclamation
        Else
            ' Loop through folder structure, creating as needed.
            strVerified = varParts(0)
            For intPart = 1 To UBound(varParts)
                strVerified = strVerified & "\" & varParts(intPart)
                If Not FSO.FolderExists(strVerified) Then FSO.CreateFolder strVerified
            Next intPart
        End If
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetDBProperty
' Author    : Adam Waller
' Date      : 9/1/2017
' Purpose   : Get a database property (Default to MDB version)
'---------------------------------------------------------------------------------------
'
Public Function GetDBProperty(strName As String) As Variant

    Dim prp As Object ' DAO.Property
    Dim oParent As Object
    
    ' Get parent container for properties
    If CurrentProject.ProjectType = acADP Then
        Set oParent = CurrentProject.Properties
    Else
        Set oParent = CurrentDb.Properties
    End If
    
    ' Look for property by name
    For Each prp In oParent
        If prp.Name = strName Then
            GetDBProperty = prp.Value
            Exit For
        End If
    Next prp
    Set prp = Nothing
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : SetDBProperty
' Author    : Adam Waller
' Date      : 9/1/2017
' Purpose   : Set a database property
'---------------------------------------------------------------------------------------
'
Public Sub SetDBProperty(ByVal strName As String, ByVal varValue As Variant, Optional ByVal prpType As Long = dbText)

    Dim prp As Object ' DAO.Property
    Dim blnFound As Boolean
    Dim dbs As DAO.Database
    Dim oParent As Object
    
    ' Properties set differently for databases and ADP projects
    If CurrentProject.ProjectType = acADP Then
        Set oParent = CurrentProject.Properties
    Else
        Set dbs = CurrentDb
        Set oParent = dbs.Properties
    End If
    
    ' Look for property in collection
    For Each prp In oParent
        If prp.Name = strName Then
            ' Check for matching type
            If Not dbs Is Nothing Then
                If prp.Type <> prpType Then
                    ' Remove so we can add it back in with the correct type.
                    dbs.Properties.Delete strName
                    Exit For
                End If
            End If
            blnFound = True
            ' Skip set on matching value
            If prp.Value = varValue Then
                Set dbs = Nothing
            Else
                ' Update value
                prp.Value = varValue
            End If
            Exit Sub
        End If
    Next prp
    
    ' Add new property
    If Not blnFound Then
        If CurrentProject.ProjectType = acADP Then
            CurrentProject.Properties.Add strName, varValue
        Else
            Set prp = dbs.CreateProperty(strName, prpType, varValue)
            dbs.Properties.Append prp
            Set dbs = Nothing
        End If
    End If
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : CloseAllFormsReports
' Author    : Adam Waller
' Date      : 1/25/2019
' Purpose   : Close all open forms and reports. Returns true if successful.
'---------------------------------------------------------------------------------------
'
Public Function CloseAllFormsReports() As Boolean

    Dim strName As String
    Dim intOpened As Integer
    Dim intItem As Integer
    
    ' Get count of opened objects
    intOpened = Forms.Count + Reports.Count
    If intOpened > 0 Then
        On Error GoTo ErrorHandler
        ' Loop through forms
        For intItem = Forms.Count - 1 To 0 Step -1
            If Forms(intItem).Caption <> "MSAccessVCS" Then
                DoCmd.Close acForm, Forms(intItem).Name
                DoEvents
            End If
            intOpened = intOpened - 1
        Next intItem
        ' Loop through reports
        Do While Reports.Count > 0
            strName = Reports(0).Name
            DoCmd.Close acReport, strName
            DoEvents
            intOpened = intOpened - 1
        Loop
        If intOpened = 0 Then CloseAllFormsReports = True
    Else
        ' No forms or reports currently open.
        CloseAllFormsReports = True
    End If
    
    Exit Function

ErrorHandler:
    Debug.Print "Error closing " & strName & ": " & Err.Number & vbCrLf & Err.Description
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetVBEExtByType
' Author    : Adam Waller
' Date      : 6/2/2015
' Purpose   : Return a standardized VBE component extension by type
'---------------------------------------------------------------------------------------
'
Public Function GetVBEExtByType(cmp As VBComponent) As String
    Dim strExt As String
    Select Case cmp.Type
        Case vbext_ct_StdModule:    strExt = ".bas"
        Case vbext_ct_MSForm:       strExt = ".frm" ' (not used in Microsoft Access)
        Case Else ' vbext_ct_Document, vbext_ct_ActiveXDesigner, vbext_ct_ClassModule
            strExt = ".cls"
    End Select
    GetVBEExtByType = strExt
End Function


'---------------------------------------------------------------------------------------
' Procedure : Shell2
' Author    : Adam Waller
' Date      : 6/3/2015
' Purpose   : Alternative to VBA Shell command, to work around issues with the
'           : TortoiseSVN command line for commits.
'---------------------------------------------------------------------------------------
'
Public Sub Shell2(strCmd As String)
    Dim objShell As WshShell
    Set objShell = New WshShell
    objShell.Exec strCmd
    Set objShell = Nothing
End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetSafeFileName
' Author    : Adam Waller
' Date      : 1/14/2019
' Purpose   : Replace illegal filename characters with URL encoded substitutes
'           : Sources: http://stackoverflow.com/questions/1976007/what-characters-are-forbidden-in-windows-and-linux-directory-names
'---------------------------------------------------------------------------------------
'
Public Function GetSafeFileName(strName As String) As String

    Dim strSafe As String

    ' Use URL encoding for these characters
    ' https://www.w3schools.com/tags/ref_urlencode.asp
    strSafe = Replace(strName, "%", "%25")  ' This should be done first.
    strSafe = Replace(strSafe, "<", "%3C")
    strSafe = Replace(strSafe, ">", "%3E")
    strSafe = Replace(strSafe, ":", "%3A")
    strSafe = Replace(strSafe, """", "%22")
    strSafe = Replace(strSafe, "/", "%2F")
    strSafe = Replace(strSafe, "\", "%5C")
    strSafe = Replace(strSafe, "|", "%7C")
    strSafe = Replace(strSafe, "?", "%3F")
    strSafe = Replace(strSafe, "*", "%2A")

    ' Return the sanitized file name.
    GetSafeFileName = strSafe
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetObjectNameFromFileName
' Author    : Adam Waller
' Date      : 5/6/2020
' Purpose   : Return the object name after translating the HTML encoding back to normal
'           : file name characters.
'---------------------------------------------------------------------------------------
'
Public Function GetObjectNameFromFileName(strFile As String) As String

    Dim strName As String
    
    strName = FSO.GetBaseName(strFile)
    ' Make sure the following list matches the one above.
    strName = Replace(strName, "%3C", "<")
    strName = Replace(strName, "%3E", ">")
    strName = Replace(strName, "%3A", ":")
    strName = Replace(strName, "%22", """")
    strName = Replace(strName, "%2F", "/")
    strName = Replace(strName, "%5C", "\")
    strName = Replace(strName, "%7C", "|")
    strName = Replace(strName, "%3F", "?")
    strName = Replace(strName, "%2A", "*")
    strName = Replace(strName, "%25", "%")  ' This should be done last.
    
    ' Return the object name
    GetObjectNameFromFileName = strName
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : HasMoreRecentChanges
' Author    : Adam Waller
' Date      : 4/27/2020
' Purpose   : Returns true if the database object has been modified more recently
'           : than the exported file or source object.
'---------------------------------------------------------------------------------------
'
Public Function HasMoreRecentChanges(objItem As IDbComponent) As Boolean
    ' File dates could be a second off (between exporting the file and saving the report)
    ' so ignore changes that are less than three seconds apart.
    If objItem.DateModified > 0 And objItem.SourceModified > 0 Then
        HasMoreRecentChanges = (DateDiff("s", objItem.DateModified, objItem.SourceModified) < -3)
    Else
        ' If we can't determine one or both of the dates, return true so the
        ' item is processed as though more recent changes were detected.
        HasMoreRecentChanges = True
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : PreserveModificationStatusBeforeCompact
' Author    : Adam Waller
' Date      : 10/13/2017
' Purpose   : Run this function before doing a compact/repair on the database to allow
'           : quick saves (skipping unchanged files) even after the dates are reset
'           : by the compact/repair operation.
'---------------------------------------------------------------------------------------
'
Public Sub PreserveModificationStatusBeforeCompact()

    Dim colContainers As New Collection
    Dim obj As Object
    Dim dbs As Database
    Dim strValue As String
    Dim varContainer As Variant
    Dim dteOldest As Date
    Dim dteCreated As Date
    Dim dteModified As Date
    Dim blnExport As Boolean
    Dim dteLastCompact As Date
    
    ' Start with today and work backwards
    dteOldest = Now
    
    ' Get date/time when the database was last compacted/repaired.
    strValue = GetDBProperty("InitiatedCompactRepair")
    If IsDate(strValue) Then dteLastCompact = CDate(strValue)
    
    ' Add object types to collection
    With colContainers
        If CurrentProject.ProjectType = acMDB Then
            Set dbs = CurrentDb
            .Add Forms
            .Add Reports
            '.Add dbs.QueryDefs
            '.Add dbs.TableDefs
            '.Add CurrentProject.AllMacros
        Else
            .Add CurrentProject.AllForms
            .Add CurrentProject.AllReports
            '.Add CurrentProject.AllMacros
        End If
    End With
    
    ' Go through each container
    For Each varContainer In colContainers
    
        ' Loop through each object
        For Each obj In varContainer
        
            ' Get creation and modified dates
            dteCreated = obj.DateCreated
            dteModified = obj.DateModified
            
            ' Default to needing to export the current object.
            blnExport = True
            
            ' If dates match, the object has not changed since last compact/repair
            If DatesClose(dteCreated, dteModified) And DatesClose(dteCreated, dteLastCompact, 20) Then
                ' Sounds like this object has not changed.
            Else
                ' Changes were made since creation or modification.
                ' Increment flag to force update on next export.
                If GetChangeFlag(obj, 0) = 0 Then SetChangeFlag obj, 1
            End If
            
            'Debug.Print obj. & " ";  & " " & obj.Name
        Next obj
        
    Next varContainer

    ' Save the current time at the database level
    SetDBProperty "InitiatedCompactRepair", CStr(Now)
    
    ' Clean up
    Set obj = Nothing
    Set colContainers = Nothing
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetChangeFlag
' Author    : Adam Waller
' Date      : 1/25/2019
' Purpose   : Get or set the custom change flag in a database object.
'---------------------------------------------------------------------------------------
'
Public Function GetChangeFlag(obj As AccessObject, Optional intDefault As Integer) As Integer
    Dim strValue As String
    strValue = GetAccessObjectProperty(obj, "GitLabChangeFlag", CStr(intDefault))
    If IsNumeric(strValue) Then GetChangeFlag = CInt(strValue)
End Function
Public Sub SetChangeFlag(obj As AccessObject, varValue As Variant)
    SetAccessObjectProperty obj, "GitLabChangeFlag", CStr(varValue)
End Sub


'---------------------------------------------------------------------------------------
' Procedure : DatesClose
' Author    : Adam Waller
' Date      : 10/13/2017
' Purpose   : Returns true if the dates are within the threshhold.
'           : (Used when dates are very similar, but not exact)
'---------------------------------------------------------------------------------------
'
Public Function DatesClose(dte1 As Date, dte2 As Date, Optional lngMaxDiffSeconds As Long = 3) As Boolean
    DatesClose = (Abs(DateDiff("s", dte1, dte2)) < lngMaxDiffSeconds)
End Function


'---------------------------------------------------------------------------------------
' Procedure : SetAccessObjectProperty
' Author    : Adam Waller
' Date      : 10/13/2017
' Purpose   : Sets a custom access object property.
'---------------------------------------------------------------------------------------
'
Public Sub SetAccessObjectProperty(objItem As AccessObject, strProperty As String, strValue As String)
    Dim prp As AccessObjectProperty
    For Each prp In objItem.Properties
        If StrComp(prp.Name, strProperty, vbTextCompare) = 0 Then
            ' Update value of property.
            prp.Value = strValue
            Exit Sub
        End If
    Next prp
    ' Property not found. Create it.
    objItem.Properties.Add strProperty, strValue
End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetAccessObjectProperty
' Author    : Adam Waller
' Date      : 10/13/2017
' Purpose   : Get the value of a custom access property
'---------------------------------------------------------------------------------------
'
Public Function GetAccessObjectProperty(objItem As AccessObject, strProperty As String, Optional strDefault As String) As Variant
    Dim prp As AccessObjectProperty
    For Each prp In objItem.Properties
        If StrComp(prp.Name, strProperty, vbTextCompare) = 0 Then
            GetAccessObjectProperty = prp.Value
            Exit Function
        End If
    Next prp
    ' Nothing found. Return default
    GetAccessObjectProperty = strDefault
End Function


'---------------------------------------------------------------------------------------
' Procedure : StripDboPrefix
' Author    : Adam Waller
' Date      : 1/25/2019
' Purpose   : Removes the dbo prefix, as sometimes encountered with ADP projects
'           : depending on the sql permissions of the current user.
'---------------------------------------------------------------------------------------
'
Public Function StripDboPrefix(strName As String) As String
    If Left$(strName, 4) = "dbo." Then
        StripDboPrefix = Mid$(strName, 5)
    Else
        StripDboPrefix = strName
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : MultiReplace
' Author    : Adam Waller
' Date      : 1/18/2019
' Purpose   : Does a string replacement of multiple items in one call.
'---------------------------------------------------------------------------------------
'
Public Function MultiReplace(ByVal strText As String, ParamArray varPairs()) As String
    Dim intPair As Integer
    For intPair = 0 To UBound(varPairs) Step 2
        strText = Replace(strText, varPairs(intPair), varPairs(intPair + 1))
    Next intPair
    MultiReplace = strText
End Function


'---------------------------------------------------------------------------------------
' Procedure : ShowIDE
' Author    : Adam Waller
' Date      : 5/18/2015
' Purpose   : Show the VBA code editor (used in autoexec macro)
'---------------------------------------------------------------------------------------
'
Public Function ShowIDE() As Boolean
    DoCmd.RunCommand acCmdVisualBasicEditor
    DoEvents
    ShowIDE = True
End Function


'---------------------------------------------------------------------------------------
' Procedure : ProgramFilesFolder
' Author    : Adam Waller
' Date      : 5/15/2015
' Purpose   : Returns the program files folder on the OS. (32 or 64 bit)
'---------------------------------------------------------------------------------------
'
Public Function ProgramFilesFolder() As String
    Dim strFolder As String
    strFolder = Environ$("PROGRAMFILES")
    ' Should always work, but just in case!
    If strFolder = vbNullString Then strFolder = "C:\Program Files (x86)"
    ProgramFilesFolder = strFolder & "\"
End Function


'---------------------------------------------------------------------------------------
' Procedure : ProjectIsSelected
' Author    : Adam Waller
' Date      : 5/15/2015
' Purpose   : Returns true if the base project is selected in the VBE
'---------------------------------------------------------------------------------------
'
Public Function ProjectIsSelected() As Boolean
    ProjectIsSelected = (Application.VBE.SelectedVBComponent Is Nothing)
End Function


'---------------------------------------------------------------------------------------
' Procedure : SelectionInActiveProject
' Author    : Adam Waller
' Date      : 5/15/2015
' Purpose   : Returns true if the current selection is in the active project
'---------------------------------------------------------------------------------------
'
Public Function SelectionInActiveProject() As Boolean
    SelectionInActiveProject = (Application.VBE.ActiveVBProject.FileName = GetUncPath(CurrentProject.FullName))
End Function


'---------------------------------------------------------------------------------------
' Procedure : IsLoaded
' Author    : Adam Waller
' Date      : 9/22/2017
' Purpose   : Returns true if the object is loaded and not in design view.
'---------------------------------------------------------------------------------------
'
Public Function IsLoaded(intType As AcObjectType, strName As String, Optional blnAllowDesignView As Boolean = False) As Boolean

    Dim frm As Form
    Dim ctl As Control
    
    If SysCmd(acSysCmdGetObjectState, intType, strName) <> adStateClosed Then
        If blnAllowDesignView Then
            IsLoaded = True
        Else
            Select Case intType
                Case acReport
                    IsLoaded = Reports(strName).CurrentView <> acCurViewDesign
                Case acForm
                    IsLoaded = Forms(strName).CurrentView <> acCurViewDesign
                Case acServerView
                    IsLoaded = CurrentData.AllViews(strName).CurrentView <> acCurViewDesign
                Case acStoredProcedure
                    IsLoaded = CurrentData.AllStoredProcedures(strName).CurrentView <> acCurViewDesign
                Case Else
                    ' Other unsupported object
                    IsLoaded = True
            End Select
        End If
    Else
        ' Could be loaded as subform
        If intType = acForm Then
            For Each frm In Forms
                For Each ctl In frm.Controls
                    If TypeOf ctl Is SubForm Then
                        If ctl.SourceObject = strName Then
                            IsLoaded = True
                            Exit For
                        End If
                    End If
                Next ctl
                If IsLoaded Then Exit For
            Next frm
        End If
    End If
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : MsgBox2
' Author    : Adam Waller
' Date      : 1/27/2017
' Purpose   : Alternate message box with bold prompt on first line.
'---------------------------------------------------------------------------------------
'
Public Function MsgBox2(strBold As String, Optional strLine1 As String, Optional strLine2 As String, Optional intButtons As VbMsgBoxStyle = vbOKOnly, Optional strTitle As String) As VbMsgBoxResult
    
    Dim strMsg As String
    Dim varLines(0 To 3) As String
    
    ' Escape single quotes by doubling them.
    varLines(0) = Replace(strBold, "'", "''")
    varLines(1) = Replace(strLine1, "'", "''")
    varLines(2) = Replace(strLine2, "'", "''")
    varLines(3) = Replace(strTitle, "'", "''")
    
    If varLines(3) = vbNullString Then varLines(3) = Application.VBE.ActiveVBProject.Name
    strMsg = "MsgBox('" & varLines(0) & "@" & varLines(1) & "@" & varLines(2) & "@'," & intButtons & ",'" & varLines(3) & "')"
    MsgBox2 = Eval(strMsg)
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : LoadOptions
' Author    : Adam Waller
' Date      : 4/15/2020
' Purpose   : Loads the current options from defaults and this project.
'---------------------------------------------------------------------------------------
'
Public Function LoadOptions() As clsOptions
    Dim Options As clsOptions
    Set Options = New clsOptions
    Options.LoadProjectOptions
    Set LoadOptions = Options
End Function


'---------------------------------------------------------------------------------------
' Procedure : Options
' Author    : Adam Waller
' Date      : 5/2/2020
' Purpose   : A global property to access options from anywhere in code.
'           : (Avoiding a global state is better OO programming, but this approach keeps
'           :  the coding simpler when you don't have to tie everything back to the
'           :  primary object.) I.e. You can just use `Encrypt("text")` instead of
'           :  having to use `Options.Encrypt("text")`
'           : To clear the current set of options, simply set the property to nothing.
'---------------------------------------------------------------------------------------
'
Public Property Get Options() As clsOptions
    If m_Options Is Nothing Then Set m_Options = LoadOptions
    Set Options = m_Options
End Property
Public Property Set Options(cNewOptions As clsOptions)
    Set m_Options = cNewOptions
End Property


'---------------------------------------------------------------------------------------
' Procedure : GetVCSVersion
' Author    : Adam Waller
' Date      : 1/28/2019
' Purpose   : Gets the version of the version control system. (Used to turn off fast
'           : save until a full export has been run with the current version of
'           : the MSAccessVCS addin.)
'---------------------------------------------------------------------------------------
'
Public Function GetVCSVersion() As String

    Dim dbs As Database
    Dim prp As DAO.Property

    Set dbs = CodeDb

    For Each prp In dbs.Properties
        If prp.Name = "AppVersion" Then
            ' Return version
            GetVCSVersion = prp.Value
        End If
    Next prp

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetVBProjectForCurrentDB
' Author    : Adam Waller
' Date      : 7/25/2017
' Purpose   : Get the actual VBE project for the current top-level database.
'           : (This is harder than you would think!)
'---------------------------------------------------------------------------------------
'
Public Function GetVBProjectForCurrentDB() As VBProject
    Set GetVBProjectForCurrentDB = GetProjectByName(CurrentProject.FullName)
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetCodeVBProject
' Author    : Adam Waller
' Date      : 4/24/2020
' Purpose   : Get a reference to the VB Project for the running code.
'---------------------------------------------------------------------------------------
'
Public Function GetCodeVBProject() As VBProject
    Set GetCodeVBProject = GetProjectByName(CodeProject.FullName)
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetProjectByName
' Author    : Adam Waller
' Date      : 5/26/2020
' Purpose   : Return the VBProject by file path. (Also supports network drives)
'---------------------------------------------------------------------------------------
'
Private Function GetProjectByName(ByVal strPath As String) As VBProject

    Dim objProj As VBIDE.VBProject
    Dim strUncPath As String
    
    ' Use currently active project by default
    Set GetProjectByName = VBE.ActiveVBProject
    
    ' VBProject filenames are UNC paths
    strUncPath = GetUncPath(strPath)
    
    If VBE.ActiveVBProject.FileName <> strUncPath Then
        ' Search for project with matching filename.
        For Each objProj In VBE.VBProjects
            If objProj.FileName = strUncPath Then
                Set GetProjectByName = objProj
                Exit For
            End If
        Next objProj
    End If
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : RunInCurrentProject
' Author    : Adam Waller
' Date      : 4/22/2020
' Purpose   : Use the Run command but make sure it is running in the context of the
'           : current project, not the add-in file.
'---------------------------------------------------------------------------------------
'
Public Sub RunSubInCurrentProject(strSubName As String)

    Dim strCmd As String
    
    ' Don't need the parentheses after the sub name
    strCmd = Replace(strSubName, "()", vbNullString)
    
    ' Make sure we are not trying to run a function with arguments
    If InStr(strCmd, "(") > 0 Then
        MsgBox2 "Unable to Run Command", _
            "Parameters are not supported for this command.", _
            "If you need to use parameters, please create a wrapper sub or function with" & vbCrLf & _
            "no parameters that you can call instead of " & strSubName & ".", vbExclamation
        Exit Sub
    End If
    
    ' Add project name so we can run it from the current datbase
    strCmd = "[" & GetVBProjectForCurrentDB.Name & "]." & strCmd
    
    ' Run the sub
    Application.Run strCmd

End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetFilePathsInFolder
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Returns a collection containing the full paths of files in a folder.
'           : Wildcards are supported.
'---------------------------------------------------------------------------------------
'
Public Function GetFilePathsInFolder(strFolder As String, Optional strFilePattern As String = "*.*") As Collection
    
    Dim oFile As Scripting.File
    Dim strBaseFolder As String
    
    strBaseFolder = StripSlash(strFolder)
    Set GetFilePathsInFolder = New Collection
    
    If FSO.FolderExists(strBaseFolder) Then
        For Each oFile In FSO.GetFolder(strBaseFolder).Files
            ' Add files that match the pattern.
            If oFile.Name Like strFilePattern Then GetFilePathsInFolder.Add oFile.Path
        Next oFile
    End If
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetSubfolderPaths
' Author    : Adam Waller
' Date      : 7/30/2020
' Purpose   : Return a collection of subfolders inside a folder.
'---------------------------------------------------------------------------------------
'
Public Function GetSubfolderPaths(strPath As String) As Collection

    Dim strBase As String
    Dim oFolder As Scripting.Folder
    
    Set GetSubfolderPaths = New Collection
    
    strBase = StripSlash(strPath)
    If FSO.FolderExists(strBase) Then
        For Each oFolder In FSO.GetFolder(strBase).SubFolders
            GetSubfolderPaths.Add oFolder.Path
        Next oFolder
    End If
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : WriteJsonFile
' Author    : Adam Waller
' Date      : 4/24/2020
' Purpose   : Creates a json file with an info header giving some clues about the
'           : contents of the file. (Helps with upgrades or changes later.)
'           : Set blnIgnoreHeaderOnlyChanges to true when the file should only be
'           : written when the dItems dictionary value changes. This helps reduce the
'           : number of files marked as changed when the actual content is the same,
'           : but a newer version of VCS was used to export the file.
'---------------------------------------------------------------------------------------
'
Public Sub WriteJsonFile(ClassMe As Object, dItems As Dictionary, strFile As String, strDescription As String, _
    Optional blnIgnoreHeaderOnlyChanges As Boolean = True)
    
    Dim dContents As Dictionary
    Dim dHeader As Dictionary
    Dim dFile As Dictionary
    Dim dExisting As Dictionary
    
    Set dContents = New Dictionary
    Set dHeader = New Dictionary
    
    ' Compare with existing file
    If blnIgnoreHeaderOnlyChanges Then
        If FSO.FileExists(strFile) Then
            Set dFile = ReadJsonFile(strFile)
            If Not dFile Is Nothing Then
                If dFile.Exists("Items") Then
                    Set dExisting = dFile("Items")
                    If DictionaryEqual(dItems, dExisting) Then
                        ' No changes to content. Leave existing file.
                        Exit Sub
                    End If
                End If
            End If
        End If
    End If
    
    ' Build dictionary structure
    dHeader.Add "Class", TypeName(ClassMe)
    dHeader.Add "Description", strDescription
    dHeader.Add "VCS Version", GetVCSVersion
    dContents.Add "Info", dHeader
    dContents.Add "Items", dItems
    
    ' Write to file in Json format
    WriteFile ConvertToJson(dContents, JSON_WHITESPACE), strFile
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ReadJsonFile
' Author    : Adam Waller
' Date      : 5/5/2020
' Purpose   : Reads a Json file into a dictionary object
'---------------------------------------------------------------------------------------
'
Public Function ReadJsonFile(strPath As String) As Dictionary
    
    Dim strText As String
    Dim stm As ADODB.Stream
    
    If FSO.FileExists(strPath) Then
        Set stm = New ADODB.Stream
        With stm
            .Charset = "UTF-8"
            .Open
            .LoadFromFile strPath
            strText = .ReadText(adReadAll)
            .Close
        End With
        
        ' If it looks like json content, then parse into a dictionary object.
        If Left$(strText, 3) = UTF8_BOM Then strText = Mid$(strText, 4)
        If Left$(strText, 1) = "{" Then Set ReadJsonFile = ParseJson(strText)
    End If
    
    Set stm = Nothing
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetSQLObjectModifiedDate
' Author    : Adam Waller
' Date      : 10/11/2017
' Purpose   : Get the last modified date for the SQL object
'---------------------------------------------------------------------------------------
'
Public Function GetSQLObjectModifiedDate(strName As String, eType As eSqlObjectType) As Date

    ' Use static variables so we can avoid hundreds of repeated calls
    ' for the same object type. Instead use a local array after
    ' pulling the initial values.
    ' (Makes a significant performance gain in complex databases)
    Static colCache As Collection
    Static strLastType As String
    Static dteCacheDate As Date

    Dim rst As ADODB.Recordset
    Dim strSql As String
    Dim strObject As String
    Dim strTypeFilter As String
    Dim intPos As Integer
    Dim strSchema As String
    Dim varItem As Variant
    Dim strType As String
    
    ' Shortcut to clear the cached variable
    If strName = vbNullString And strType = vbNullString Then
        Set colCache = Nothing
        strLastType = vbNullString
        dteCacheDate = 0
        Exit Function
    End If
    
    ' Only try this on ADP projects
    If CurrentProject.ProjectType <> acADP Then Exit Function
    
    ' Simple validation on object name
    strObject = Replace(strName, ";", vbNullString)
    
    ' Build schema filter if required
    intPos = InStr(1, strObject, ".")
    If intPos > 0 Then
        strObject = Mid$(strObject, intPos + 1)
        strSchema = Left$(strName, intPos - 1)
        'strSchemaFilter = " AND [schema_id]=schema_id('" & strSchema & "')"
    Else
        strSchema = "dbo"
    End If
    
    ' Build type filter
    Select Case eType
        Case estView: strType = "V"
        Case estStoredProcedure: strType = "P"
        Case estTable: strType = "U"
        Case estTrigger: strType = "TR"
    End Select
    If strType <> vbNullString Then strTypeFilter = " AND [type]='" & strType & "'"
    
    ' Check to see if we have already cached the results
    If strType = strLastType And (DateDiff("s", dteCacheDate, Now) < 5) And Not colCache Is Nothing Then
        ' Look through cache to find matching date
        For Each varItem In colCache
            If varItem(0) = strName Then
                GetSQLObjectModifiedDate = varItem(1)
                Exit For
            End If
        Next varItem
    Else
        ' Look up from query, and cache results
        Set colCache = New Collection
        dteCacheDate = Now
        strLastType = strType
        
        ' Build SQL query to find object
        strSql = "SELECT [name], schema_name([schema_id]) as [schema], modify_date FROM sys.objects WHERE 1=1 " & strTypeFilter
        Set rst = New ADODB.Recordset
        With rst
            .Open strSql, CurrentProject.Connection, adOpenForwardOnly, adLockReadOnly
            Do While Not .EOF
                ' Return date when name matches. (But continue caching additional results)
                If Nz(!Name) = strObject And Nz(!schema) = strSchema Then GetSQLObjectModifiedDate = Nz(!modify_date)
                If Nz(!schema) = "dbo" Then
                    colCache.Add Array(Nz(!Name), Nz(!modify_date))
                Else
                    ' Include schema name in object name
                    colCache.Add Array(Nz(!schema) & "." & Nz(!Name), Nz(!modify_date))
                End If
                .MoveNext
            Loop
            .Close
        End With
        Set rst = Nothing
    End If
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetSQLObjectDefinitionForADP
' Author    : awaller
' Date      : 12/12/2016
' Purpose   : Returns the SQL definition for the ADP project item.
'           : (Queries, Views, Tables, etc... are not stored in Access but on the
'           :  SQL server.)
'           : NOTE: This takes a simplistic approach, which does not guard againts
'           : certain types of SQL injection attacks. Use at your own risk!
'---------------------------------------------------------------------------------------
'
Public Function GetSQLObjectDefinitionForADP(strName As String) As String
    
    Dim rst As ADODB.Recordset
    Dim strSql As String
    Dim strObject As String
    
    ' Only try this on ADP projects
    If CurrentProject.ProjectType <> acADP Then Exit Function
    
    ' Simple validation on object name
    strObject = Replace(strName, ";", vbNullString)
    
    strSql = "SELECT object_definition (OBJECT_ID(N'" & strObject & "'))"
    '@Ignore SetAssignmentWithIncompatibleObjectType
    Set rst = CurrentProject.Connection.Execute(strSql)
    If Not rst.EOF Then
        ' Get SQL definition
        GetSQLObjectDefinitionForADP = Nz(rst(0).Value)
    End If
    
    Set rst = Nothing
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : Perf
' Author    : Adam Waller
' Date      : 11/3/2020
' Purpose   : Wrapper for performance logging class
'---------------------------------------------------------------------------------------
'
Public Function Perf() As clsPerformance
    If m_Perf Is Nothing Then Set m_Perf = New clsPerformance
    Set Perf = m_Perf
End Function


'---------------------------------------------------------------------------------------
' Procedure : Log
' Author    : Adam Waller
' Date      : 4/28/2020
' Purpose   : Wrapper for log file class
'---------------------------------------------------------------------------------------
'
Public Function Log() As clsLog
    If m_Log Is Nothing Then Set m_Log = New clsLog
    Set Log = m_Log
End Function


'---------------------------------------------------------------------------------------
' Procedure : FSO
' Author    : Adam Waller
' Date      : 1/18/2019
' Purpose   : Wrapper for file system object. A property allows us to clear the object
'           : reference when we have completed an export or import operation.
'---------------------------------------------------------------------------------------
'
Public Property Get FSO() As FileSystemObject
    If m_FSO Is Nothing Then Set m_FSO = New FileSystemObject
    Set FSO = m_FSO
End Property
Public Property Set FSO(ByVal RHS As FileSystemObject)
    Set m_FSO = RHS
End Property


'---------------------------------------------------------------------------------------
' Procedure : SaveComponentAsText
' Author    : Adam Waller
' Date      : 4/29/2020
' Purpose   : Wrapper for Application.SaveAsText that verifies that the path exists,
'           : and then removes any existing file before saving the object as text.
'---------------------------------------------------------------------------------------
'
Public Sub SaveComponentAsText(intType As AcObjectType, strName As String, strFile As String)
    
    Dim strTempFile As String
    
    On Error GoTo ErrHandler
    
    ' Export to temporary file
    strTempFile = GetTempFile
    Perf.OperationStart "App.SaveAsText()"
    Application.SaveAsText intType, strName, strTempFile
    Perf.OperationEnd
    
    ' Sanitize certain object types
    Select Case intType
        Case acForm, acReport, acQuery, acMacro
            ' Sanitizing converts to UTF-8
            If FSO.FileExists(strFile) Then DeleteFile (strFile)
            SanitizeFile strTempFile
            FSO.MoveFile strTempFile, strFile
        Case Else
            ' Handle UCS conversion if needed
            ConvertUcs2Utf8 strTempFile, strFile
    End Select
    
    ' Normal exit
    On Error GoTo 0
    Exit Sub
    
ErrHandler:
    If Err.Number = 2950 And intType = acTableDataMacro Then
        ' This table apparently didn't have a Table Data Macro.
        Exit Sub
    Else
        ' Some other error.
        Err.Raise Err.Number
    End If
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : LoadComponentFromText
' Author    : Adam Waller
' Date      : 5/5/2020
' Purpose   : Load the object into the database from the saved source file.
'---------------------------------------------------------------------------------------
'
Public Sub LoadComponentFromText(intType As AcObjectType, strName As String, strFile As String)

    Dim strTempFile As String
    Dim blnConvert As Boolean
    
    ' Check UCS-2-LE requirement for the current database.
    ' (Cached after first call)
    Select Case intType
        Case acForm, acReport, acQuery, acMacro, acTableDataMacro
            blnConvert = RequiresUcs2
    End Select
    
    ' Only run conversion if needed.
    If blnConvert Then
        ' Perform file conversion, and import from temp file.
        strTempFile = GetTempFile
        ConvertUtf8Ucs2 strFile, strTempFile, False
        Perf.OperationStart "App.LoadFromText()"
        Application.LoadFromText intType, strName, strTempFile
        Perf.OperationEnd
        FSO.DeleteFile strTempFile, True
    Else
        ' Load UTF-8 file
        Perf.OperationStart "App.LoadFromText()"
        Application.LoadFromText intType, strName, strFile
        Perf.OperationEnd
    End If
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : SecureBetween
' Author    : Casper Englund
' Date      : 2020-06-03
' Purpose   : Secures content between two strings.
'---------------------------------------------------------------------------------------
'
Public Function SecureBetween(strText As String, strStartAfter As String, strEndBefore As String, Optional Compare As VbCompareMethod) As String
        
        If strText = vbNullString Or Options.Security = esNone Then
            SecureBetween = strText
        Else
            If Options.Security = esEncrypt Then
                SecureBetween = EncryptBetween(strText, strStartAfter, strEndBefore, Compare)
            ElseIf Options.Security = esRemove Then
                Dim lngPos As Long
                Dim lngStart As Long
                Dim lngLen As Long
                
                lngPos = InStr(1, strText, strStartAfter, Compare)
                If lngPos > 0 Then
                    lngStart = lngPos + Len(strStartAfter) - 1
                    lngPos = InStr(lngStart + 1, strText, strEndBefore)
                    If lngPos > 0 Then
                        lngLen = lngPos - lngStart
                    End If
                End If
                
                If lngLen = 0 Then
                    ' No tags found. Return original string
                    SecureBetween = strText
                Else
                    SecureBetween = Left$(strText, lngStart) & Mid$(strText, lngStart + lngLen)
                End If
    
            End If
        End If
        
End Function

'---------------------------------------------------------------------------------------
' Procedure : SecurePath
' Author    : Adam Waller
' Date      : 6/1/2020
' Purpose   : Secures just the folder path, not the filename.
'---------------------------------------------------------------------------------------
'
Public Function SecurePath(strPath As String) As String

    Dim strParent As String

    strParent = FSO.GetParentFolderName(strPath)
    If strParent = vbNullString Then
        ' Could be relative path or just a filename.
        SecurePath = strPath
    Else
        If Options.Security = esRemove Then
            SecurePath = FSO.GetFileName(strPath)
        Else
            ' Could be encrypted or plain text, depending on options.
            SecurePath = Secure(strParent) & "\" & FSO.GetFileName(strPath)
        End If
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : FolderHasVcsOptionsFile
' Author    : Adam Waller
' Date      : 5/5/2020
' Purpose   : Returns true if the folder as a vcs-options.json file, which is required
'           : to build a project from source files.
'---------------------------------------------------------------------------------------
'
Public Function FolderHasVcsOptionsFile(strFolder As String) As Boolean
    FolderHasVcsOptionsFile = FSO.FileExists(StripSlash(strFolder) & "\vcs-options.json")
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetOriginalDbFullPathFromSource
' Author    : Adam Waller
' Date      : 5/5/2020
' Purpose   : Determine the original full path of the database, based on the files
'           : in the source folder.
'---------------------------------------------------------------------------------------
'
Public Function GetOriginalDbFullPathFromSource(strFolder As String) As String
    
    Dim strPath As String
    Dim dContents As Dictionary
    Dim strFile As String
    
    strPath = StripSlash(strFolder) & "\vbe-project.json"
    If FSO.FileExists(strPath) Then
        Set dContents = ReadJsonFile(strPath)
        strFile = Decrypt(dNZ(dContents, "Items\FileName"))
        If Left$(strFile, 4) = "rel:" Then
            ' Use parent folder of source folder
            GetOriginalDbFullPathFromSource = StripSlash(strFolder) & "\..\" & FSO.GetFileName(Mid$(strFile, 5))
        ElseIf InStr(1, strFile, "@{") > 0 Then
            ' Decryption failed.
            ' We might be able to figure out a relative path from the export path.
            strPath = StripSlash(strFolder) & "\vcs-options.json"
            If FSO.FileExists(strPath) Then
                Set dContents = ReadJsonFile(strPath)
                ' Make sure we can read something, but that the export folder is blank.
                ' (Default, which indicates that it would be in the parent folder of the
                '  source directory.)
                If dNZ(dContents, "Info\AddinVersion") <> vbNullString _
                    And dNZ(dContents, "Options\ExportFolder") = vbNullString Then
                    ' Use parent folder of source directory
                    GetOriginalDbFullPathFromSource = StripSlash(strFolder) & "\..\" & FSO.GetFileName(strFile)
                End If
            End If
        Else
            ' Return full path to file.
            GetOriginalDbFullPathFromSource = strFile
        End If
    End If
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : dNZ
' Author    : Adam Waller
' Date      : 3/23/2020
' Purpose   : Like the NZ function but for dictionary elements
'---------------------------------------------------------------------------------------
'
Public Function dNZ(dObject As Dictionary, strPath As String, Optional strDelimiter As String = "\") As String

    Dim varPath As Variant
    Dim intCnt As Integer
    Dim dblVal As Double
    Dim strKey As String
    Dim varSegment As Variant
        
    ' Split path into parts
    varPath = Split(strPath, strDelimiter)
    Set varSegment = dObject

    For intCnt = LBound(varPath) To UBound(varPath)

        strKey = varPath(intCnt)
        If dObject Is Nothing Then
            ' No object found
            Exit For
        ElseIf TypeOf varSegment Is Collection Then
            ' Expect index (integer)
            If IsNumeric(strKey) Then
                ' Looks like an array index
                dblVal = CDbl(strKey)
                ' Do a couple more checks to see if this looks like a valid index
                If dblVal < 1 Or dblVal > 32000 Or dblVal <> CInt(dblVal) Then Exit For
                ' See if this is the last segment
                If intCnt = UBound(varPath) Then
                    If TypeOf varSegment(dblVal) Is Dictionary Then
                        ' Need a named key
                        Exit For
                    Else
                        ' Could be an array of values
                        dNZ = Nz(varSegment(dblVal))
                    End If
                Else
                    ' Move out to next segment
                    Set varSegment = varSegment(dblVal)
                End If
            End If
        ElseIf TypeOf varSegment Is Dictionary Then
            ' Expect key (string)
            If intCnt = UBound(varPath) Then
                ' Reached last segment
                If varSegment.Exists(strKey) Then
                    If TypeOf varSegment Is Dictionary Then
                        dNZ = Nz(varSegment(strKey))
                    Else
                        ' Might be array
                        Exit For
                    End If
                End If
            Else
                ' Move out to next segment
                If varSegment.Exists(strKey) And Not IsEmpty(varSegment(strKey)) Then
                    Set varSegment = varSegment(strKey)
                Else
                    ' Path not found
                    Exit For
                End If
            End If
        End If
    Next intCnt

End Function


'---------------------------------------------------------------------------------------
' Procedure : TableExists
' Author    : Adam Waller
' Date      : 5/7/2020
' Purpose   : Returns true if the table object is found in the dabase.
'---------------------------------------------------------------------------------------
'
Public Function TableExists(strName As String) As Boolean
    TableExists = Not (DCount("*", "MSysObjects", "Name=""" & strName & """ AND Type=1") = 0)
End Function


'---------------------------------------------------------------------------------------
' Procedure : SortDictionaryByKeys
' Author    : Adam Waller
' Date      : 5/8/2020
' Purpose   : Rebuilds a dictionary object by adding all the items to a new dictionary
'           : sorted by keys.
'---------------------------------------------------------------------------------------
'
Public Function SortDictionaryByKeys(dSource As Dictionary) As Dictionary

    Dim dSorted As Dictionary
    Dim varKeys() As Variant
    Dim varKey As Variant
    Dim lngCnt As Long
    
    ' Don't need to sort empty dictionary or single item
    If dSource.Count < 2 Then
        Set SortDictionaryByKeys = dSource
        Exit Function
    End If
    
    Dim sourceCount As Long
    sourceCount = dSource.Count - 1
    
    ' Build and sort array of keys
    ReDim varKeys(0 To sourceCount)
    For Each varKey In dSource.Keys
        varKeys(lngCnt) = varKey
        lngCnt = lngCnt + 1
    Next varKey
    
    QuickSort varKeys, 0, sourceCount
    
    ' Build and return new dictionary using sorted keys
    Set dSorted = New Dictionary
    For lngCnt = 0 To sourceCount
        dSorted.Add varKeys(lngCnt), dSource(varKeys(lngCnt))
    Next lngCnt
    
    Set SortDictionaryByKeys = dSorted
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : QuickSort
' Author    : Stack Overflow
' Date      : 5/8/2020
' Purpose   : Adapted from https://stackoverflow.com/a/152325/4121863
' Usage     : QuickSort MyArray
'---------------------------------------------------------------------------------------
'
Public Sub QuickSort(ByRef vArray As Variant, Optional ByVal inLow As Long, Optional ByVal inHi As Long)

    Dim pivot   As Variant
    Dim tmpSwap As Variant
    Dim tmpLow  As Long
    Dim tmpHi   As Long
    
    If IsMissing(inLow) Then inLow = LBound(vArray)
    If IsMissing(inHi) Then inHi = UBound(vArray)
    
    tmpLow = inLow
    tmpHi = inHi
    
    pivot = vArray((inLow + inHi) \ 2)
    
    While (tmpLow <= tmpHi)
        While (vArray(tmpLow) < pivot And tmpLow < inHi)
            tmpLow = tmpLow + 1
        Wend
        
        While (pivot < vArray(tmpHi) And tmpHi > inLow)
            tmpHi = tmpHi - 1
        Wend
        
        If (tmpLow <= tmpHi) Then
            tmpSwap = vArray(tmpLow)
            vArray(tmpLow) = vArray(tmpHi)
            vArray(tmpHi) = tmpSwap
            tmpLow = tmpLow + 1
            tmpHi = tmpHi - 1
        End If
    Wend
    
    If (inLow < tmpHi) Then QuickSort vArray, inLow, tmpHi
    If (tmpLow < inHi) Then QuickSort vArray, tmpLow, inHi
  
End Sub


'---------------------------------------------------------------------------------------
' Procedure : SetDAOProperty
' Author    : Adam Waller
' Date      : 5/8/2020
' Purpose   : Updates a DAO property, adding if it does not exist or is the wrong type.
'---------------------------------------------------------------------------------------
'
Public Sub SetDAOProperty(objParent As Object, intType As Integer, strName As String, varValue As Variant)

    Dim prp As DAO.Property
    Dim blnFound As Boolean
    
    ' Look through existing properties.
    For Each prp In objParent.Properties
        If prp.Name = strName Then
            blnFound = True
            Exit For
        End If
    Next prp

    ' Verify type, and update value if found.
    If blnFound Then
        If prp.Type <> intType Then
            objParent.Properties.Delete strName
            blnFound = False
        Else
            If objParent.Properties(strName).Value <> varValue Then
                objParent.Properties(strName).Value = varValue
            End If
        End If
    End If
        
    ' Add new property if needed
    If Not blnFound Then
        ' Create property, then append to collection
        Set prp = objParent.CreateProperty(strName, intType, varValue)
        objParent.Properties.Append prp
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetRelativePath
' Author    : Adam Waller
' Date      : 5/11/2020
' Purpose   : Returns a path relative to current database.
'           : If a relative path is not possible, it returns the original full path.
'---------------------------------------------------------------------------------------
'
Public Function GetRelativePath(strPath As String) As String
    
    Dim strFolder As String
    Dim strUncPath As String
    Dim strUncTest As String
    Dim strRelative As String
    
    ' Check for matching parent folder as relative to the project path.
    strFolder = GetUncPath(CurrentProject.Path) & "\"
    
    ' Default to original path if no relative path could be resolved.
    strRelative = strPath
    
    ' Compare strPath to the current project path
    If InStr(1, strPath, strFolder, vbTextCompare) = 1 Then
        ' In export folder or subfolder. Simple replacement
        strRelative = "rel:" & Mid$(strPath, Len(strFolder) + 1)
    Else
        ' Make sure we have a path, not just a file name.
        If InStr(1, strRelative, "\") > 0 Then
            ' Check UNC path for network drives
            strUncPath = GetUncPath(strPath)
            If StrComp(strUncPath, strPath, vbTextCompare) <> 0 Then
                ' We are dealing with a network drive
                strUncTest = GetRelativePath(strUncPath)
                If StrComp(strUncPath, strUncTest, vbTextCompare) <> 0 Then
                    ' Resolved to relative UNC path
                    strRelative = strUncTest
                End If
            End If
        End If
    End If
    
    ' Return relative (or original) path
    GetRelativePath = strRelative

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetPathFromRelative
' Author    : Adam Waller
' Date      : 5/11/2020
' Purpose   : Expands a relative path out to the full path.
'---------------------------------------------------------------------------------------
'
Public Function GetPathFromRelative(strPath As String) As String
    If Left$(strPath, 4) = "rel:" Then
        GetPathFromRelative = CurrentProject.Path & "\" & Mid$(strPath, 5)
    Else
        ' No relative path used.
        GetPathFromRelative = strPath
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : Secure
' Author    : Adam Waller
' Date      : 6/1/2020
' Purpose   : Secure the text based on the loaded option.
'---------------------------------------------------------------------------------------
'
Public Function Secure(strText As String) As String
    Select Case Options.Security
        Case esEncrypt: Secure = Encrypt(strText)
        Case esRemove:  Secure = vbNullString
        Case esNone:    Secure = strText
    End Select
End Function


'---------------------------------------------------------------------------------------
' Procedure : DictionaryEqual
' Author    : Adam Waller
' Date      : 6/2/2020
' Purpose   : Returns true if the two dictionary objects are equal in values to each
'           : other, including nested values. Testing the quickest comparisons first
'           : to make the function as performant as possible.
'---------------------------------------------------------------------------------------
'
Public Function DictionaryEqual(dOne As Dictionary, dTwo As Dictionary) As Boolean

    Dim strOne As String
    Dim strTwo As String
    Dim blnEqual As Boolean
    
    If dOne Is Nothing And dTwo Is Nothing Then
        ' Neither object set.
        blnEqual = True
    ElseIf Not dOne Is Nothing And Not dTwo Is Nothing Then
        ' Both are objects. Check count property.
        If dOne.Count = dTwo.Count Then
            strOne = ConvertToJson(dOne)
            strTwo = ConvertToJson(dTwo)
            ' Compare string length
            If Len(strOne) = Len(strTwo) Then
                ' Perform a binary (case-sensitive) comparison of strings.
                blnEqual = (StrComp(strOne, strTwo, vbBinaryCompare) = 0)
            End If
        End If
    End If
    
    ' Return comparison result
    DictionaryEqual = blnEqual
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : CreateZipFile
' Author    : Adam Waller
' Date      : 5/26/2020
' Purpose   : Create an empty zip file to copy files into.
'           : Adapted from: http://www.rondebruin.nl/win/s7/win001.htm
'---------------------------------------------------------------------------------------
'
Public Sub CreateZipFile(strPath As String)
    
    Dim strHeader As String
    Dim intFile As Integer
    
    ' Build Zip file header
    strHeader = "PK" & Chr$(5) & Chr$(6) & String$(18, 0)
    
    ' Write to file
    VerifyPath strPath
    With FSO.CreateTextFile(strPath, True)
        .Write strHeader
        .Close
    End With
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : CopyToZip
' Author    : Adam Waller
' Date      : 5/26/2020
' Purpose   : Copy a file into a zip archive.
'           : Adapted from: http://www.rondebruin.nl/win/s7/win001.htm
'---------------------------------------------------------------------------------------
'
Public Sub CopyFileToZip(strFile As String, strZip As String)
    
    Dim oApp As Object
    Dim varZip As Variant
    Dim varFile As Variant
    
    ' Must use variants for the CopyHere function to work.
    varZip = strZip
    varFile = strFile
    
    Set oApp = CreateObject("Shell.Application")
    oApp.Namespace(varZip).CopyHere varFile
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : CopyFolderToZip
' Author    : Adam Waller
' Date      : 6/3/2020
' Purpose   : Copies a folder of items into a zip file.
'---------------------------------------------------------------------------------------
'
Public Sub CopyFolderToZip(strFolder As String, strZip As String, _
    Optional blnPauseTillFinished As Boolean = True, Optional intTimeoutSeconds As Integer = 60)

    Dim oApp As Object
    Dim varZip As Variant
    Dim varFolder As Variant
    Dim sngTimeout As Single
    Dim lngCount As Long
    
    ' Must use variants for the CopyHere function to work.
    varZip = strZip
    varFolder = strFolder
    
    ' Count the total items before we start the copy,
    ' since there might already be files in the zip folder.
    Set oApp = CreateObject("Shell.Application")
    lngCount = oApp.Namespace(varFolder).Items.Count + oApp.Namespace(varZip).Items.Count
    
    ' Start the copy
    oApp.Namespace(varZip).CopyHere oApp.Namespace(varFolder).Items
    
    ' Pause till the copying is complete, or we hit the timeout.
    sngTimeout = Timer + intTimeoutSeconds
    Do While Timer < sngTimeout
        ' Check to see if all the items have been copied.
        If oApp.Namespace(varZip).Items.Count = lngCount Then Exit Do
        Pause 0.5
    Loop
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ExtractFromZip
' Author    : Adam Waller
' Date      : 6/3/2020
' Purpose   : Extracts all the files from a zip archive. (Requires a .zip extension)
'---------------------------------------------------------------------------------------
'
Public Sub ExtractFromZip(strZip As String, strDestFolder As String, _
    Optional blnPauseTillFinished As Boolean = True, Optional intTimeoutSeconds As Integer = 60)

    Dim oApp As Object
    Dim varZip As Variant
    Dim varFolder As Variant
    Dim sngTimeout As Single
    Dim lngCount As Long
    Dim strFolder As String
    
    ' Build folder path, and make sure it exists
    If Not FSO.FolderExists(strDestFolder) Then FSO.CreateFolder strDestFolder
    strFolder = FSO.GetFolder(strDestFolder).Path
    
    ' Must use variants for the CopyHere function to work.
    varZip = strZip
    varFolder = strFolder & "\"

    ' Count the total items before we start the copy,
    ' since there might already be files in the zip folder.
    Set oApp = CreateObject("Shell.Application")
    If blnPauseTillFinished Then
        lngCount = oApp.Namespace(varFolder).Items.Count + oApp.Namespace(varZip).Items.Count
    End If

    ' Begin the extraction
    oApp.Namespace(varFolder).CopyHere oApp.Namespace(varZip).Items
    If blnPauseTillFinished Then
        ' Pause till the copying is complete, or we hit the timeout.
        sngTimeout = Timer + intTimeoutSeconds
        Do While Timer < sngTimeout
            ' Check to see if all the items have been copied.
            If oApp.Namespace(varZip).Items.Count = lngCount Then Exit Do
            Pause 0.5
        Loop
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Pause
' Author    : Adam Waller
' Date      : 6/3/2020
' Purpose   : Pause the code execution for x seconds.
'---------------------------------------------------------------------------------------
'
Public Sub Pause(sngSeconds As Single)
    Sleep sngSeconds * 1000
End Sub


'---------------------------------------------------------------------------------------
' Procedure : DatabaseOpen
' Author    : Adam Waller
' Date      : 7/14/2020
' Purpose   : Returns true if a database (or ADP project) is currently open.
'---------------------------------------------------------------------------------------
'
Public Function DatabaseOpen() As Boolean
    DatabaseOpen = Not (CurrentDb Is Nothing And CurrentProject.Connection Is Nothing)
    'DatabaseOpen = Workspaces(0).Databases.Count > 0   ' Another approach
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetUncPath
' Author    : Adam Waller
' Date      : 7/14/2020
' Purpose   : Returns the UNC path for a network location (if applicable)
'---------------------------------------------------------------------------------------
'
Public Function GetUncPath(strPath As String)

    Dim strDrive As String
    Dim strUNC As String
    
    strUNC = strPath
    strDrive = FSO.GetDriveName(strPath)
    With FSO.GetDrive(strDrive)
        If .DriveType = Remote Then
            strUNC = Replace(strPath, strDrive, .ShareName, , 1, vbTextCompare)
        End If
    End With
    GetUncPath = strUNC
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : CheckForLegacyModules
' Author    : Adam Waller
' Date      : 7/16/2020
' Purpose   : Informs the user if the database contains a legacy module from another
'           : fork of this project. (Some users might not realize that these are not
'           : needed anymore.)
'---------------------------------------------------------------------------------------
'
Public Sub CheckForLegacyModules()
    If FSO.FileExists(Options.GetExportFolder & "modules\VCS_ImportExport.bas") Then
        MsgBox2 "Legacy Files not Needed", _
            "Other forks of the MSAccessVCS project used additional VBA modules to export code." & vbCrLf & _
            "This is no longer needed when using the installed Version Control Add-in.", _
            "Feel free to remove the legacy VCS_* modules from your database project and enjoy" & vbCrLf & _
            "a simpler, cleaner code base for ongoing development.  :-)", vbInformation, "Just a Suggestion..."
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetLastModifiedDate
' Author    : Adam Waller
' Date      : 7/30/2020
' Purpose   : Get the last modified date on a folder or file with Unicode support.
'---------------------------------------------------------------------------------------
'
Public Function GetLastModifiedDate(strPath As String) As Date
    
    Dim oFile As Scripting.File
    Dim oFolder As Scripting.Folder
    
    If FSO.FileExists(strPath) Then
        Set oFile = FSO.GetFile(strPath)
        GetLastModifiedDate = oFile.DateLastModified
    ElseIf FSO.FolderExists(strPath) Then
        Set oFolder = FSO.GetFolder(strPath)
        GetLastModifiedDate = oFolder.DateLastModified
    End If
        
End Function