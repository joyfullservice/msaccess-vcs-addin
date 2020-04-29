Option Explicit
Option Compare Database
Option Private Module

Public Const JSON_WHITESPACE As String = "    "

Public colVerifiedPaths As New Collection

' Formats used when exporting table data.
Public Enum eTableDataExportFormat
    etdNoData = 0
    etdTabDelimited = 1
    etdXML = 2
    [_last] = 2
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
    edbGalleryImage
    edbDocumentObject
    edbSavedSpec
    edbNavPaneGroups
    edbVbeProject
    edbVbeReference
End Enum


' Logging class
Private m_Log As clsLog

' Keep a persistent reference to file system object after initializing version control.
' This way we don't have to recreate this object dozens of times while using VCS.
Private m_FSO As Scripting.FileSystemObject


'---------------------------------------------------------------------------------------
' Procedure : SanitizeFile
' Author    : Adam Waller
' Date      : 1/23/2019
' Purpose   : Sanitize the text file (forms and reports)
'---------------------------------------------------------------------------------------
'
Public Sub SanitizeFile(strPath As String, cOptions As clsOptions)

    Dim sngOverall As Single
    Dim sngTimer As Single
    Dim cData As New clsConcat
    Dim strText As String
    Dim rxBlock As New VBScript_RegExp_55.RegExp
    Dim rxLine As New VBScript_RegExp_55.RegExp
    Dim rxIndent As New VBScript_RegExp_55.RegExp
    Dim objMatches As VBScript_RegExp_55.MatchCollection
    Dim blnIsReport As Boolean
    Dim cPattern As New clsConcat
    Dim stmInFile As Scripting.TextStream
    Dim blnGetLine As Boolean
    
    On Error GoTo 0
    
    ' Timers to monitor performance
    sngTimer = Timer
    sngOverall = sngTimer
        
    '  Setup Block matching Regex.
    rxBlock.IgnoreCase = False
    
    ' Build main search patterns
    With cPattern
    
        '  Match PrtDevNames / Mode with or without W
        If cOptions.AggressiveSanitize Then .Add "(?:"
        .Add "PrtDev(?:Names|Mode)[W]?"
        If cOptions.AggressiveSanitize Then
          '  Add and group aggressive matches
          .Add "|GUID|""GUID""|NameMap|dbLongBinary ""DOL"""
          .Add ")"
        End If
        
        '  Ensure that this is the begining of a block.
        .Add " = Begin"
        
        ' Set block search pattern
        rxBlock.Pattern = .GetStr
        .Clear
        
        '  Setup Line Matching Regex.
        .Add "^\s*(?:"
        .Add "Checksum ="
        .Add "|BaseInfo|NoSaveCTIWhenDisabled =1"
        If cOptions.StripPublishOption Then
            .Add "|dbByte ""PublishToWeb"" =""1"""
            .Add "|PublishOption =1"
        End If
        .Add ")"

        ' Set line search pattern
        rxLine.Pattern = .GetStr
    End With
    
    ' Open file to read contents line by line.
    Set stmInFile = FSO.OpenTextFile(strPath, ForReading)

    blnGetLine = True
    Do Until stmInFile.AtEndOfStream
    
        ' Only call DoEvents once per second.
        ' (Drastic performance gains)
        If Timer - sngTimer > 1 Then
            DoEvents
            sngTimer = Timer
        End If
    
        ' Check if we need to get a new line of text
        If blnGetLine Then
            strText = stmInFile.ReadLine
        Else
            blnGetLine = True
        End If
        
        ' Skip lines starting with line pattern
        If rxLine.Test(strText) Then
            
            ' set up initial pattern
            rxIndent.Pattern = "^(\s+)\S"
            
            ' Get indentation level.
            Set objMatches = rxIndent.Execute(strText)
            
            ' Setup pattern to match current indent
            Select Case objMatches.Count
                Case 0
                    rxIndent.Pattern = "^" & vbNullString
                Case Else
                    rxIndent.Pattern = "^" & objMatches(0).SubMatches(0)
            End Select
            rxIndent.Pattern = rxIndent.Pattern & "\S"
            
            ' Skip lines with deeper indentation
            Do While Not stmInFile.AtEndOfStream
                strText = stmInFile.ReadLine
                If rxIndent.Test(strText) Then Exit Do
            Loop
            
            ' We've moved on at least one line so restart the
            ' regex testing when starting the loop again.
            blnGetLine = False
        
        ' Skip blocks of code matching block pattern
        ElseIf rxBlock.Test(strText) Then
            Do While Not stmInFile.AtEndOfStream
                strText = stmInFile.ReadLine
                If InStr(strText, "End") Then Exit Do
            Loop
        
        ' Check for report object
        ElseIf InStr(1, strText, "Begin Report") = 1 Then
            blnIsReport = True
            cData.Add strText
            cData.Add vbCrLf
            
        ' Watch for end of report (and skip these lines)
        ElseIf blnIsReport And (InStr(1, strText, "    Right =") Or InStr(1, strText, "    Bottom =")) Then
            If InStr(1, strText, "    Bottom =") Then blnIsReport = False

        ' Regular lines of data to add.
        Else
            cData.Add strText
            cData.Add vbCrLf
        End If
        
    Loop
    
    ' Close and delete original file
    stmInFile.Close
    FSO.DeleteFile strPath
    
    ' Write file all at once, rather than line by line.
    ' (Otherwise the code can bog down with tens of thousands of write operations)
    WriteFile cData.GetStr, strPath

    ' Show stats if debug turned on.
    Log.Add "    Sanitized in " & Format(Timer - sngOverall, "0.00") & " seconds.", cOptions.ShowDebug

End Sub


'---------------------------------------------------------------------------------------
' Procedure : SanitizeXML
' Author    : Adam Waller
' Date      : 4/27/2020
' Purpose   : Remove non-essential data that changes every time the file is exported.
'---------------------------------------------------------------------------------------
'
Public Sub SanitizeXML(strPath As String, cOptions As clsOptions)

    Dim sngOverall As Single
    Dim sngTimer As Single
    Dim cData As clsConcat
    Dim strText As String
    Dim rxLine As VBScript_RegExp_55.RegExp
    Dim objMatches As VBScript_RegExp_55.MatchCollection
    Dim blnIsReport As Boolean
    Dim stmInFile As Scripting.TextStream
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
    'rxLine.Pattern = "^\s*(?:<dataroot xmlns:(.+))( generated="".+"")"
    
    ' Open file to read contents line by line.
    Set stmInFile = FSO.OpenTextFile(strPath, ForReading)

    ' Loop through all the lines in the file
    Do Until stmInFile.AtEndOfStream
        
        ' Read line from file
        strText = stmInFile.ReadLine
                 
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
    Log.Add "    Sanitized in " & Format(Timer - sngOverall, "0.00") & " seconds.", cOptions.ShowDebug

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
    If Right(ProjectPath, 1) <> "\" Then ProjectPath = ProjectPath & "\"
End Function


'---------------------------------------------------------------------------------------
' Procedure : MkDirIfNotExist
' Author    : Adam Waller
' Date      : 1/25/2019
' Purpose   : Create folder `Path`. Silently do nothing if it already exists.
'---------------------------------------------------------------------------------------
'
Public Sub MkDirIfNotExist(strPath As String)
    If Not FSO.FolderExists(StripSlash(strPath)) Then MkDir StripSlash(strPath)
End Sub


'---------------------------------------------------------------------------------------
' Procedure : clearfilesbyextension
' Author    : Adam Waller
' Date      : 1/25/2019
' Purpose   : Erase all *.`ext` files in `Path`.
'---------------------------------------------------------------------------------------
'
Public Sub ClearFilesByExtension(ByVal strFolder As String, strExt As String)
    If Not FSO.FolderExists(StripSlash(strFolder)) Then Exit Sub
    If Dir(strFolder & "*." & strExt) <> "" Then
        FSO.DeleteFile strFolder & "*." & strExt
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ClearTextFilesForFastSave
' Author    : Adam Waller
' Date      : 12/14/2016
' Purpose   : Clears existing source files that don't have a matching object in the
'           : database.
'---------------------------------------------------------------------------------------
'
Public Sub ClearOrphanedSourceFiles(cType As IDbComponent, ParamArray StrExtensions())
    
    Dim oFolder As Scripting.Folder
    Dim oFile As Scripting.File
    Dim colNames As Collection
    Dim strFile As String
    Dim varName As Variant
    Dim varExt As Variant
    Dim strPrimaryExt As String
    
    ' No orphaned files if the folder doesn't exist.
    If Not FSO.FolderExists(cType.BaseFolder) Then Exit Sub
    
    ' Cache a list of source file names
    Set colNames = cType.GetFileList
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
                'If strFile = "modsavedSpecs.bas" Then Stop
                ' Remove any file that doesn't have a matching name.
                If Not InCollection(colNames, strFile) Then
                    ' Object not found in database. Remove file.
                    Kill oFile.ParentFolder.Path & "\" & oFile.Name
                    Log.Add "  Removing orphaned file: " & strFile, cType.Options.ShowDebug
                End If
                
                ' No need to check other extensions since we
                ' already had a match and processed the file.
                Exit For
            End If
        Next varExt
    Next oFile
    
    ' Remove base folder if we don't have any files in it
    If oFolder.Files.Count = 0 Then oFolder.Delete
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : StripSlash
' Author    : Adam Waller
' Date      : 1/25/2019
' Purpose   : Strip the trailing slash
'---------------------------------------------------------------------------------------
'
Public Function StripSlash(strText As String) As String
    If Right(strText, 1) = "\" Then
        StripSlash = Left(strText, Len(strText) - 1)
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
Public Function InCollection(MyCol As Collection, MyValue) As Boolean
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
' Date      : 5/15/2015
' Purpose   : Verifies that the path to a folder exists, caching results to
'           : avoid uneeded calls to the Dir() function.
'---------------------------------------------------------------------------------------
'
Public Sub VerifyPath(strFolderPath As String)
    
    Dim varPath As Variant
    
    ' Check cache first
    For Each varPath In colVerifiedPaths
        If strFolderPath = varPath Then
            ' Found path. Assume it still exists
            Exit Sub
        End If
    Next varPath
    
    ' If code reaches here, we don't have a copy of the path
    ' in the cached list of verified paths. Verify and add
    If Dir(strFolderPath, vbDirectory) = "" Then
        ' Path does not seem to exist. Create it.
        MkDirIfNotExist strFolderPath
    End If
    colVerifiedPaths.Add strFolderPath
    
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
Public Sub SetDBProperty(strName As String, varValue, Optional prpType = dbText)

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
    Dim frm As Form
    Dim rpt As Report
    
    ' Get count of opened objects
    intOpened = Forms.Count + Reports.Count
    If intOpened > 0 Then
        On Error GoTo ErrorHandler
        ' Loop through forms
        For intItem = Forms.Count - 1 To 0 Step -1
            If Forms(intItem).Name <> "frmMain" Then
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
    Dim objShell As Object
    Set objShell = CreateObject("WScript.Shell")
    objShell.Exec strCmd
    Set objShell = Nothing
End Sub


'---------------------------------------------------------------------------------------
' Procedure : WriteFile
' Author    : Adam Waller
' Date      : 1/23/2019
' Purpose   : Save string variable to text file. (Building the folder path if needed)
'---------------------------------------------------------------------------------------
'
Public Sub WriteFile(strContent As String, strPath As String)

    Dim stm As New ADODB.Stream
    
    ' Make sure the path exists before we write a file.
    VerifyPath FSO.GetParentFolderName(strPath)
    
    With stm
        ' Use Unicode file encoding if needed.
        If StringHasUnicode(strContent) Then
            .Charset = "utf-8"
        Else
            ' Use ASCII text.
            .Charset = "us-ascii"
        End If
        .Open
        .WriteText strContent
        .SaveToFile strPath, adSaveCreateOverWrite
        .Close
    End With
    Set stm = Nothing
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : StringHasUnicode
' Author    : Adam Waller
' Date      : 3/6/2020
' Purpose   : Returns true if the string contains non-ASCI characters.
'---------------------------------------------------------------------------------------
'
Public Function StringHasUnicode(strText As String) As Boolean
    Dim reg As New VBScript_RegExp_55.RegExp
    With reg
        .Pattern = "[^\u0000-\u007F]"
        StringHasUnicode = .Test(strText)
    End With
End Function


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
    strSafe = Replace(strName, "%", "%25")  ' Since we are using this character for encoding. (Makes decoding easier if we do that at some point in the future.)
    strSafe = Replace(strName, "<", "%3C")
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
Public Function GetAccessObjectProperty(objItem As AccessObject, strProperty As String, Optional strDefault As String)
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
    If Left(strName, 4) = "dbo." Then
        StripDboPrefix = Mid(strName, 5)
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
Public Function ShowIDE()
    DoCmd.RunCommand acCmdVisualBasicEditor
    DoEvents
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
    If strFolder = "" Then strFolder = "C:\Program Files (x86)"
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
    SelectionInActiveProject = (Application.VBE.ActiveVBProject.FileName = UncPath(CurrentProject.FullName))
End Function


'---------------------------------------------------------------------------------------
' Procedure : UncPath
' Author    : Adam Waller
' Date      : 5/18/2015
' Purpose   : Returns the UNC path of a mapped network drive, if applicable
'---------------------------------------------------------------------------------------
'
Public Function UncPath(strPath As String) As String
    
    Dim strDrive As String
    Dim strShare As String
    
    ' Identify drive letter and share name
    With FSO
        strDrive = .GetDriveName(.GetAbsolutePathName(strPath))
        strShare = .GetDrive(strDrive).ShareName
    End With
    
    If strShare <> "" Then
        ' Replace drive with UNC path
        UncPath = strShare & Mid(strPath, Len(strDrive) + 1)
    Else
        ' Return unmodified path
        UncPath = strPath
    End If
        
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
    
    If varLines(3) = "" Then varLines(3) = Application.VBE.ActiveVBProject.Name
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
    Dim cOptions As clsOptions
    Set cOptions = New clsOptions
    cOptions.LoadDefaultOptions
    cOptions.LoadProjectOptions
    Set LoadOptions = cOptions
End Function


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
' Procedure : TimerIcon
' Author    : Adam Waller
' Date      : 4/16/2020
' Purpose   : Return the next increment of a timer icon, updating no more than a half
'           : second between increments.
'           : https://emojipedia.org/search/?q=clock
'---------------------------------------------------------------------------------------
'
Public Function TimerIcon() As String
    
    Static intHour As Integer
    Static sngLast As Single
    
    Dim strClocks As String
    
    ' Build list of clock characters
    ' (Need to figure out the AscW value for the clock characters)
    
    If (Timer - sngLast > 0.5) Or (Timer < sngLast) Then
        If intHour = 12 Then intHour = 0
        intHour = intHour + 1
    End If
    
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

    Dim objProj As Object
    Dim strPath As String
    
    strPath = CurrentProject.FullName
    If VBE.ActiveVBProject.FileName = strPath Then
        ' Use currently active project
        Set GetVBProjectForCurrentDB = VBE.ActiveVBProject
    Else
        ' Search for project with matching filename.
        For Each objProj In VBE.VBProjects
            If objProj.FileName = strPath Then
                Set GetVBProjectForCurrentDB = objProj
                Exit For
            End If
        Next objProj
    End If
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetCodeVBProject
' Author    : Adam Waller
' Date      : 4/24/2020
' Purpose   : Get a reference to the VB Project for the running code.
'---------------------------------------------------------------------------------------
'
Public Function GetCodeVBProject() As VBProject

    Dim objProj As VBIDE.VBProject
    Dim strPath As String
    
    strPath = CodeProject.FullName
    If VBE.ActiveVBProject.FileName = strPath Then
        ' Use currently active project
        Set GetCodeVBProject = VBE.ActiveVBProject
    Else
        ' Search for project with matching filename.
        For Each objProj In VBE.VBProjects
            If objProj.FileName = strPath Then
                Set GetCodeVBProject = objProj
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
    strCmd = Replace(strSubName, "()", "")
    
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
'---------------------------------------------------------------------------------------
'
Public Function GetFilePathsInFolder(strDirPath As String, Optional Attributes As VbFileAttribute = vbNormal) As Collection
    
    Dim strBaseFolder As String
    Dim strFile As String
    
    strBaseFolder = FSO.GetParentFolderName(strDirPath) & "\"
    Set GetFilePathsInFolder = New Collection
    strFile = Dir(strDirPath, Attributes)
    Do While strFile <> vbNullString
        GetFilePathsInFolder.Add strFile
        strFile = Dir()
    Loop
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : WriteJsonFile
' Author    : Adam Waller
' Date      : 4/24/2020
' Purpose   : Creates a json file with an info header giving some clues about the
'           : contents of the file. (Helps with upgrades or changes later.)
'---------------------------------------------------------------------------------------
'
Public Sub WriteJsonFile(ClassMe As Object, dItems As Scripting.Dictionary, strFile As String, strDescription As String)
    
    Dim dContents As Scripting.Dictionary
    Dim dHeader As Scripting.Dictionary
    
    Set dContents = New Scripting.Dictionary
    Set dHeader = New Scripting.Dictionary
    
    ' Build dictionary structure
    dHeader.Add "Class", TypeName(ClassMe)
    dHeader.Add "Description", strDescription
    dHeader.Add "VCS Version", GetVCSVersion
    dContents.Add "Info", dHeader
    dContents.Add "Items", dItems
    
    ' Write to file in Json format
    WriteFile ConvertToJson(dContents, JSON_WHITESPACE) & vbCrLf, strFile
    
End Sub


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
    Dim strSQL As String
    Dim strObject As String
    Dim strTypeFilter As String
    Dim intPos As Integer
    Dim strSchema As String
    Dim strSchemaFilter As String
    Dim varItem As Variant
    Dim strType As String
    
    ' Shortcut to clear the cached variable
    If strName = "" And strType = "" Then
        Set colCache = Nothing
        strLastType = ""
        dteCacheDate = 0
        Exit Function
    End If
    
    ' Only try this on ADP projects
    If CurrentProject.ProjectType <> acADP Then Exit Function
    
    ' Simple validation on object name
    strObject = Replace(strName, ";", "")
    
    ' Build schema filter if required
    intPos = InStr(1, strObject, ".")
    If intPos > 0 Then
        strObject = Mid(strObject, intPos + 1)
        strSchema = Left(strName, intPos - 1)
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
    If strType = strLastType And (DateDiff("s", dteCacheDate, Now()) < 5) And Not colCache Is Nothing Then
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
        dteCacheDate = Now()
        strLastType = strType
        
        ' Build SQL query to find object
        strSQL = "SELECT [name], schema_name([schema_id]) as [schema], modify_date FROM sys.objects WHERE 1=1 " & strTypeFilter
        Set rst = New ADODB.Recordset
        With rst
            .Open strSQL, CurrentProject.Connection, adOpenForwardOnly, adLockReadOnly
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
    Dim strSQL As String
    Dim strObject As String
    
    ' Only try this on ADP projects
    If CurrentProject.ProjectType <> acADP Then Exit Function
    
    ' Simple validation on object name
    strObject = Replace(strName, ";", "")
    
    strSQL = "SELECT object_definition (OBJECT_ID(N'" & strObject & "'))"
    Set rst = CurrentProject.Connection.Execute(strSQL)
    If Not rst.EOF Then
        ' Get SQL definition
        GetSQLObjectDefinitionForADP = Nz(rst(0).Value)
    End If
    
    Set rst = Nothing
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : Log
' Author    : Adam Waller
' Date      : 4/28/2020
' Purpose   :
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
Public Property Get FSO() As Scripting.FileSystemObject
    If m_FSO Is Nothing Then Set m_FSO = New Scripting.FileSystemObject
    Set FSO = m_FSO
End Property
Public Property Set FSO(ByVal RHS As Scripting.FileSystemObject)
    Set m_FSO = RHS
End Property