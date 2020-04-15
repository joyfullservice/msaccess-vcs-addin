Option Explicit
Option Compare Database
Option Private Module

Public colVerifiedPaths As New Collection


'---------------------------------------------------------------------------------------
' Procedure : SanitizeFile
' Author    : Adam Waller
' Date      : 1/23/2019
' Purpose   : Sanitize the text file (forms and reports)
'---------------------------------------------------------------------------------------
'
Public Sub SanitizeFile(strPath As String, cModel As IVersionControl)

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
    rxBlock.ignoreCase = False
    
    ' Build main search patterns
    With cPattern
    
        '  Match PrtDevNames / Mode with or  without W
        If cModel.AggressiveSanitize Then .Add "(?:"
        .Add "PrtDev(?:Names|Mode)[W]?"
        If cModel.AggressiveSanitize Then
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
        If cModel.StripPublishOption Then
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
    cModel.Log "    Sanitized in " & Format(Timer - sngOverall, "0.00") & " seconds.", cModel.ShowDebug

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
' Procedure : ClearTextFilesFromDir
' Author    : Adam Waller
' Date      : 1/25/2019
' Purpose   : Erase all *.`ext` files in `Path`.
'---------------------------------------------------------------------------------------
'
Public Sub ClearTextFilesFromDir(ByVal strFolder As String, strExt As String)
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
Public Sub ClearOrphanedSourceFiles(ByVal strPath As String, objContainer As Object, cModel As IVersionControl, ParamArray StrExtensions())
    
    Dim oFolder As Scripting.Folder
    Dim oFile As Scripting.File
    Dim colNames As New Collection
    Dim objItem As Object
    Dim strFile As String
    Dim varName As Variant
    Dim blnFound As Boolean
    Dim varExt As Variant
    
    ' Continue with more in-depth review, clearing any file that
    ' is not represented by a database object.
    If Not FSO.FolderExists(strPath) Then Exit Sub

    ' Build list of database objects
    If Not objContainer Is Nothing Then
        For Each objItem In objContainer
            If TypeOf objItem Is Relation Then
                ' Exclude specific names
                Select Case objItem.Name
                    Case "MSysNavPaneGroupsMSysNavPaneGroupToObjects", "MSysNavPaneGroupCategoriesMSysNavPaneGroups"
                        ' Skip these built-in relationships
                    Case Else
                        ' Relationship names can't be used directly as file names.
                        colNames.Add GetRelationFileName(objItem)
                End Select
            Else
                colNames.Add GetSafeFileName(StripDboPrefix(objItem.Name))
            End If
        Next objItem
    End If
    
    ' Loop through files in folder
    strPath = Left(strPath, Len(strPath) - 1)
    Set oFolder = FSO.GetFolder(strPath)
    
    For Each oFile In oFolder.Files
        ' Check against list of extensions
        For Each varExt In StrExtensions
        
            ' Check for matching extension on wanted list.
            If FSO.GetExtensionName(oFile.Path) = varExt Then
                
                ' Get file name
                strFile = FSO.GetBaseName(oFile.Name)
                
                ' Loop through list of names to see if this one exists
                blnFound = False
                For Each varName In colNames
                    If strFile = varName Then
                        blnFound = True
                        Exit For
                    End If
                Next varName
                
                If Not blnFound Then
                    ' Object not found in database. Remove file.
                    Kill oFile.ParentFolder.Path & "\" & oFile.Name
                    cModel.Log "  Removing orphaned file: " & strFile, cModel.ShowDebug
                End If
                
                ' No need to check other extensions since we
                ' already had a match and processed the file.
                Exit For
            End If
        Next varExt
    Next oFile
    
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
' Procedure : PadRight
' Author    : Adam Waller
' Date      : 1/25/2019
' Purpose   : Pad a string on the right to make it `count` characters long.
'---------------------------------------------------------------------------------------
'
Public Function PadRight(strText As String, intCharacters As Integer)
    If Len(strText) < intCharacters Then
        PadRight = strText & Space(intCharacters - Len(strText))
    Else
        PadRight = strText
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
' Procedure : CloseAllFormsReports
' Author    : Adam Waller
' Date      : 1/25/2019
' Purpose   : Close all open forms and reports. Returns true if successful.
'---------------------------------------------------------------------------------------
'
Public Function CloseAllFormsReports() As Boolean

    Dim strName As String
    Dim intOpened As Integer
    
    ' Get count of opened objects
    intOpened = Forms.Count + Reports.Count
    If intOpened > 0 Then
        On Error GoTo ErrorHandler
        Do While Forms.Count > 0
            strName = Forms(0).Name
            DoCmd.Close acForm, strName
            DoEvents
        Loop
        Do While Reports.Count > 0
            strName = Reports(0).Name
            DoCmd.Close acReport, strName
            DoEvents
        Loop
        If (Forms.Count + Reports.Count) = 0 Then CloseAllFormsReports = True
        
        ' Switch back to IDE window
        ShowIDE
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
' Purpose   : Save string variable to text file.
'---------------------------------------------------------------------------------------
'
Public Sub WriteFile(strContent As String, strPath As String)
    Dim stm As New ADODB.Stream
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
' Date      : 12/14/2016
' Purpose   : Returns true if the database object has been modified more recently
'           : than the exported file.
'---------------------------------------------------------------------------------------
'
Public Function HasMoreRecentChanges(objItem As Object, strFile As String) As Boolean
    ' File dates could be a second off (between exporting the file and saving the report)
    ' so ignore changes that are less than three seconds apart.
    If Dir(strFile) <> "" Then
        HasMoreRecentChanges = (DateDiff("s", objItem.DateModified, FileDateTime(strFile)) < -3)
    Else
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


' returns substring between e.g. "(" and ")", internal brackets ar skippped
'Public Function SubString(P As Integer, s As String, startsWith As String, endsWith As String)
'    Dim start As Integer
'    Dim cursor As Integer
'    Dim p1 As Integer
'    Dim p2 As Integer
'    Dim level As Integer
'    start = InStr(P, s, startsWith)
'    level = 1
'    p1 = InStr(start + 1, s, startsWith)
'    p2 = InStr(start + 1, s, endsWith)
'    While level > 0
'        If p1 > p2 And p2 > 0 Then
'            cursor = p2
'            level = level - 1
'        ElseIf p2 > p1 And p1 > 0 Then
'            cursor = p1
'            level = level + 1
'        ElseIf p2 > 0 And p1 = 0 Then
'            cursor = p2
'            level = level - 1
'        ElseIf p1 > 0 And p1 = 0 Then
'            cursor = p1
'            level = level + 1
'        ElseIf p1 = 0 And p2 = 0 Then
'            SubString = ""
'            Exit Function
'        End If
'        p1 = InStr(cursor + 1, s, startsWith)
'        p2 = InStr(cursor + 1, s, endsWith)
'    Wend
'    SubString = Mid(s, start + 1, cursor - start - 1)
'End Function
'


Public Sub TestOptions()
    
    Dim cOpt As New clsOptions
    'cOpt.PrintOptionsToDebugWindow
    cOpt.SaveOptionsForProject
    cOpt.LoadProjectOptions
    cOpt.PrintOptionsToDebugWindow
    
End Sub



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