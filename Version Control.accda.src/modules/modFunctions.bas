Option Explicit
Option Compare Database
Option Private Module

Public ShowDebugInfo As Boolean ' Public to this project only, when used with Option Private Module
Public colVerifiedPaths As New Collection

' Constants for Scripting.FileSystemObject API
Public Enum eOpenType
    ForReading = 1
    ForWriting = 2
    ForAppending = 8
End Enum

Public Enum eTristate
    TristateTrue = -1
    TristateFalse = 0
    TristateUseDefault = -2
End Enum

Private Const AggressiveSanitize = True
Private Const StripPublishOption = True


Private m_SourcePath As String


' Can we export without closing the form?

' Export a database object with optional UCS2-to-UTF-8 conversion.
Public Sub ExportObject(obj_type_num As Integer, obj_name As String, file_path As String, cModel As IVersionControl, Optional Ucs2Convert As Boolean = False)
        
    Dim blnSkip As Boolean

    On Error GoTo ErrHandler
    
    modFunctions.VerifyPath Left(file_path, InStrRev(file_path, "\"))
    
    ' Check for fast save
    If Not cModel Is Nothing Then
        If cModel.FastSave Then
            Select Case obj_type_num
                Case acQuery
                    blnSkip = (HasMoreRecentChanges(CurrentData.AllQueries(obj_name), file_path))
                Case acForm
                    blnSkip = (HasMoreRecentChanges(CurrentProject.AllForms(obj_name), file_path))
                Case acReport
                    blnSkip = (HasMoreRecentChanges(CurrentProject.AllReports(obj_name), file_path))
                ' Tables are done through a different function
            End Select
        End If
    End If
    
    If blnSkip Then
        If cModel.ShowDebug Then Debug.Print "  (Skipping '" & obj_name & "')"
    Else
        If Ucs2Convert Then
            Dim tempFileName As String: tempFileName = modFileAccess.TempFile()
            Application.SaveAsText obj_type_num, obj_name, tempFileName
            modFileAccess.ConvertUcs2Utf8 tempFileName, file_path
        Else
            Application.SaveAsText obj_type_num, obj_name, file_path
        End If
        If cModel.ShowDebug Then Debug.Print "  " & obj_name
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
    
    If Not modFunctions.FileExists(file_path) Then Exit Sub
    
    If Ucs2Convert Then
        Dim tempFileName As String: tempFileName = modFileAccess.TempFile()
        modFileAccess.ConvertUtf8Ucs2 file_path, tempFileName
        Application.LoadFromText obj_type_num, obj_name, tempFileName
        
        Dim fso As New Scripting.FileSystemObject
        fso.DeleteFile tempFileName
    Else
        Application.LoadFromText obj_type_num, obj_name, file_path
    End If
End Sub


' For each *.txt in `Path`, find and remove a number of problematic but
' unnecessary lines of VB code that are inserted automatically by the
' Access GUI and change often (we don't want these lines of code in
' version control).
Public Sub SanitizeTextFiles(Path As String, ext As String, cModel As IVersionControl)

    Dim fileName As String
    fileName = Dir(Path & "*." & ext)
    Do Until Len(fileName) = 0
        SanitizeFile Path, fileName, ext, cModel
        fileName = Dir()
    Loop

End Sub

Public Sub QuickTest()
    Dim cModel As IVersionControl
    Set cModel = New clsModelGitHub
    cModel.ShowDebug = True
    cModel.ExportBaseFolder = CurrentProject.Path & "\" & CurrentProject.Name & ".src\"
    SanitizeFile "C:\Users\awaller.IAA\Documents\GitHub\alert-contacts\ALERTContacts.adp.src\reports\", "rptEventsReceipt.bas", "bas", cModel
End Sub


' Sanitize the text file
Public Sub SanitizeFile(strPath As String, strFile As String, strExt As String, cModel As IVersionControl)

    Dim fso As Scripting.FileSystemObject
    Dim strData As String
    Dim sngOverall As Single
    Dim sngTimer As Single
    Dim cData As New clsConcat
    Dim txt As String
        
    ' Timers to monitor performance
    sngTimer = Timer
    sngOverall = sngTimer
    
    Set fso = New Scripting.FileSystemObject
    Dim isReport As Boolean: isReport = False
    
    '  Setup Block matching Regex.
    Dim rxBlock As New RegExp
    'Set rxBlock = CreateObject("VBScript.RegExp")
    rxBlock.ignoreCase = False
    '
    '  Match PrtDevNames / Mode with or  without W
    Dim srchPattern As String
    srchPattern = "PrtDev(?:Names|Mode)[W]?"
    If (AggressiveSanitize = True) Then
      '  Add and group aggressive matches
      srchPattern = "(?:" & srchPattern
      srchPattern = srchPattern & "|GUID|""GUID""|NameMap|dbLongBinary ""DOL"""
      srchPattern = srchPattern & ")"
    End If
    '  Ensure that this is the begining of a block.
    srchPattern = srchPattern & " = Begin"
'Debug.Print srchPattern
    rxBlock.Pattern = srchPattern
    '
    '  Setup Line Matching Regex.
    Dim rxLine As New RegExp    ' Object
    'Set rxLine = CreateObject("VBScript.RegExp")
    srchPattern = "^\s*(?:"
    srchPattern = srchPattern & "Checksum ="
    srchPattern = srchPattern & "|BaseInfo|NoSaveCTIWhenDisabled =1"
    If (StripPublishOption = True) Then
        srchPattern = srchPattern & "|dbByte ""PublishToWeb"" =""1"""
        srchPattern = srchPattern & "|PublishOption =1"
    End If
    srchPattern = srchPattern & ")"
'Debug.Print srchPattern
    rxLine.Pattern = srchPattern


    Dim obj_name As String
    obj_name = Mid(strFile, 1, InStrRev(strFile, ".") - 1)

    Dim InFile As Scripting.TextStream ' Object
    Set InFile = fso.OpenTextFile(strPath & obj_name & "." & strExt, ForReading)
    Dim OutFile As Scripting.TextStream ' Object
    Set OutFile = fso.CreateTextFile(strPath & obj_name & ".sanitize", True)

    Dim getLine As Boolean: getLine = True
    Do Until InFile.AtEndOfStream
    
        ' Only call DoEvents once per second.
        ' (Drastic performance gains)
        If Timer - sngTimer > 1 Then
            DoEvents
            sngTimer = Timer
        End If
    
        ' Check if we need to get a new line of text
        If getLine = True Then
            txt = InFile.ReadLine
        Else
            getLine = True
        End If
        
        ' Skip lines starting with line pattern
        If rxLine.test(txt) Then
            Dim rxIndent As New RegExp ' Object
            'Set rxIndent = CreateObject("VBScript.RegExp")
            rxIndent.Pattern = "^(\s+)\S"
            '
            ' Get indentation level.
            Dim matches As VBScript_RegExp_55.MatchCollection ' Object
            Set matches = rxIndent.Execute(txt)
            '
            ' Setup pattern to match current indent
            Select Case matches.Count
                Case 0
                    rxIndent.Pattern = "^" & vbNullString
                Case Else
                    rxIndent.Pattern = "^" & matches(0).SubMatches(0)
            End Select
            rxIndent.Pattern = rxIndent.Pattern + "\S"
            '
            ' Skip lines with deeper indentation
            Do Until InFile.AtEndOfStream
                txt = InFile.ReadLine
                If rxIndent.test(txt) Then Exit Do
            Loop
            ' We've moved on at least one line so do get a new one
            ' when starting the loop again.
            getLine = False
        '
        ' skip blocks of code matching block pattern
        ElseIf rxBlock.test(txt) Then
            Do Until InFile.AtEndOfStream
                txt = InFile.ReadLine
                If InStr(txt, "End") Then Exit Do
            Loop
        ElseIf InStr(1, txt, "Begin Report") = 1 Then
            isReport = True
            cData.Add txt
            cData.Add vbCrLf
            'OutFile.WriteLine txt
        ElseIf isReport = True And (InStr(1, txt, "    Right =") Or InStr(1, txt, "    Bottom =")) Then
            'skip line
            If InStr(1, txt, "    Bottom =") Then
                isReport = False
            End If
        Else
            cData.Add txt
            cData.Add vbCrLf
            'OutFile.WriteLine txt
        End If
    
    Loop
    
    ' Write file all at once, rather than line by line.
    ' (Otherwise the code can bog down with tens of thousands of write operations)
    OutFile.Write cData.GetStr
    
    OutFile.Close
    InFile.Close

    ' Show stats if debug turned on.
    If cModel.ShowDebug Then Debug.Print "    Sanitized in " & Format(Timer - sngOverall, "0.00") & " seconds."

    fso.DeleteFile (strPath & strFile)
    DoEvents
    Dim thisFile As Scripting.File
    Set thisFile = fso.GetFile(strPath & obj_name & ".sanitize")
    thisFile.Move (strPath & strFile)

End Sub


' Path/Directory of the current database file.
Public Function ProjectPath() As String
    ProjectPath = CurrentProject.Path
    If Right(ProjectPath, 1) <> "\" Then ProjectPath = ProjectPath & "\"
End Function


'---------------------------------------------------------------------------------------
' Procedure : VCSSourcePath
' Author    : Adam Waller
' Date      : 5/18/2015
' Purpose   : Get source path. (Allow user to specify this)
'---------------------------------------------------------------------------------------
'
Public Property Get zVCSSourcePath() As String
    If m_SourcePath = "" Then m_SourcePath = ProjectPath & CurrentProject.Name & ".src\"
    VCSSourcePath = m_SourcePath
End Property


'---------------------------------------------------------------------------------------
' Procedure : VCSSourcePath
' Author    : Adam Waller
' Date      : 5/18/2015
' Purpose   : Set the source path for import/export
'           : (Set to "" to use default path)
'---------------------------------------------------------------------------------------
'
Public Property Let VCSSourcePath(strPath As String)
    If Len(strPath) > 0 Then
        ' Ensure we have a trailing slash
        If Right(strPath, 1) <> "\" Then strPath = strPath & "\"
    End If
    m_SourcePath = strPath
End Property


' Create folder `Path`. Silently do nothing if it already exists.
Public Sub MkDirIfNotExist(Path As String)
    On Error GoTo MkDirIfNotexist_noop
    MkDir Path
MkDirIfNotexist_noop:
    On Error GoTo 0
End Sub


' Delete a file if it exists.
Public Sub DelIfExist(Path As String)
    On Error GoTo DelIfNotExist_Noop
    Kill Path
DelIfNotExist_Noop:
    On Error GoTo 0
End Sub


' Erase all *.`ext` files in `Path`.
Public Sub ClearTextFilesFromDir(Path As String, ext As String)
    Dim fso As New Scripting.FileSystemObject
    If Not fso.FolderExists(Path) Then Exit Sub

    On Error GoTo ClearTextFilesFromDir_noop
    If Dir(Path & "*." & ext) <> "" Then
        fso.DeleteFile Path & "*." & ext
    End If
ClearTextFilesFromDir_noop:

    On Error GoTo 0
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ClearTextFilesForFastSave
' Author    : Adam Waller
' Date      : 12/14/2016
' Purpose   : Like ClearTextFilesFromDir, but only clears files that don't exist.
'           : The existing files will be compared to the database objects to avoid
'           : unecessary processing when no changes have occured.
'---------------------------------------------------------------------------------------
'
Public Sub ClearTextFilesForFastSave(ByVal strPath As String, strExt As String, strType As String)
    
    Dim fso As New Scripting.FileSystemObject
    Dim objContainer As Object
    Dim oFolder As Scripting.Folder
    Dim oFile As Scripting.File
    Dim colNames As New Collection
    Dim objItem As Object
    Dim strFile As String
    Dim varName As Variant
    Dim blnFound As Boolean
    
    If CurrentProject.ProjectType = acMDB Then
        ' Access Database
        Select Case strType
            Case "forms": Set objContainer = CurrentProject.AllForms
            Case "reports": Set objContainer = CurrentProject.AllReports
            Case "queries": Set objContainer = CurrentData.AllQueries
            Case "tables": Set objContainer = CurrentData.AllTables
            Case "macros": Set objContainer = CurrentProject.AllMacros
        Case Else
            ' Fast save not (yet) supported
            ClearTextFilesFromDir strPath, strExt
            Exit Sub
        End Select
    Else
        ' ADP Project
        Select Case strType
            Case "forms": Set objContainer = CurrentProject.AllForms
            Case "reports": Set objContainer = CurrentProject.AllReports
            Case "tables": Set objContainer = CurrentData.AllTables
            Case "macros": Set objContainer = CurrentProject.AllMacros
            Case "views": Set objContainer = CurrentData.AllViews
            Case "procedures": Set objContainer = CurrentData.AllStoredProcedures
            Case "functions": Set objContainer = CurrentData.AllFunctions
            'Case "triggers": Set objContainer = CurrentProject.All
        Case Else
            ' Fast save not (yet) supported
            ClearTextFilesFromDir strPath, strExt
            Exit Sub
        End Select
    End If
    
    
    ' Continue with more in-depth review, clearing any file that
    ' is not represented by a database object.
    If Not DirExists(strPath) Then Exit Sub

    ' Build list of database objects
    For Each objItem In objContainer
        colNames.Add GetSafeFileName(StripDboPrefix(objItem.Name))
    Next objItem
    
    ' Loop through files in folder
    strPath = Left(strPath, Len(strPath) - 1)
    Set oFolder = fso.GetFolder(strPath)
    
    For Each oFile In oFolder.Files
        If fso.GetExtensionName(oFile.Path) = strExt Then
            
            ' Get file name
            strFile = fso.GetBaseName(oFile.Name)
            
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
            End If
            
        End If
    Next oFile
    
    On Error GoTo 0
    Set fso = Nothing

    Exit Sub


ErrHandler:
    Err.Clear
    Resume Next
    
End Sub

Function DirExists(strPath As String) As Boolean
    On Error Resume Next
    DirExists = False
    DirExists = ((GetAttr(strPath) And vbDirectory) = vbDirectory)
End Function

Function FileExists(strPath As String) As Boolean
    On Error Resume Next
    FileExists = False
    FileExists = ((GetAttr(strPath) And vbDirectory) <> vbDirectory)
End Function

'--------------------
' String Functions: String Builder,String Padding (right only), Substrings
'--------------------

' String builder: Init
Public Function Sb_Init() As String()
    Dim X(-1 To -1) As String
    Sb_Init = X
End Function

' String builder: Clear
Public Sub Sb_Clear(ByRef sb() As String)
    ReDim sb(-1 To -1)
End Sub

' String builder: Append
Public Sub Sb_Append(ByRef sb() As String, Value As String)
    If LBound(sb) = -1 Then
        ReDim sb(0 To 0)
    Else
        ReDim Preserve sb(0 To UBound(sb) + 1)
    End If
    sb(UBound(sb)) = Value
End Sub

' String builder: Get value
Public Function Sb_Get(ByRef sb() As String) As String
    Sb_Get = Join(sb, "")
End Function


' Pad a string on the right to make it `count` characters long.
Public Function PadRight(Value As String, Count As Integer)
    PadRight = Value
    If Len(Value) < Count Then
        PadRight = PadRight & Space(Count - Len(Value))
    End If
End Function

' returns substring between e.g. "(" and ")", internal brackets ar skippped
Public Function SubString(P As Integer, s As String, startsWith As String, endsWith As String)
    Dim start As Integer
    Dim cursor As Integer
    Dim p1 As Integer
    Dim p2 As Integer
    Dim level As Integer
    start = InStr(P, s, startsWith)
    level = 1
    p1 = InStr(start + 1, s, startsWith)
    p2 = InStr(start + 1, s, endsWith)
    While level > 0
        If p1 > p2 And p2 > 0 Then
            cursor = p2
            level = level - 1
        ElseIf p2 > p1 And p1 > 0 Then
            cursor = p1
            level = level + 1
        ElseIf p2 > 0 And p1 = 0 Then
            cursor = p2
            level = level - 1
        ElseIf p1 > 0 And p1 = 0 Then
            cursor = p1
            level = level + 1
        ElseIf p1 = 0 And p2 = 0 Then
            SubString = ""
            Exit Function
        End If
        p1 = InStr(cursor + 1, s, startsWith)
        p2 = InStr(cursor + 1, s, endsWith)
    Wend
    SubString = Mid(s, start + 1, cursor - start - 1)
End Function


'---------------------------------------------------------------------------------------
' Procedure : InArray
' Author    : Adam Waller
' Date      : 5/14/2015
' Purpose   : Returns true if the item is found in the array
'---------------------------------------------------------------------------------------
'
Public Function InArray(varArray As Variant, varItem As Variant) As Boolean
    Dim intCnt As Integer
    If IsMissing(varArray) Then Exit Function
    If Not IsArray(varArray) Then
        InArray = (varItem = varArray)
    Else
        For intCnt = LBound(varArray) To UBound(varArray)
            If varArray(intCnt) = varItem Then
                InArray = True
                Exit For
            End If
        Next intCnt
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
' Procedure : ArrayToCollection
' Author    : Adam Waller
' Date      : 5/14/2015
' Purpose   : Convert the array to a collection
'---------------------------------------------------------------------------------------
'
Public Function ArrayToCollection(varArray As Variant) As Collection
    Dim intCnt As Integer
    Dim colItems As New Collection
    If IsArray(varArray) Then
        For intCnt = LBound(varArray) To UBound(varArray)
            colItems.Add varArray(intCnt)
        Next intCnt
    End If
    Set ArrayToCollection = colItems
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


' Close all open forms.
Public Function CloseFormsReports()
    On Error GoTo errorHandler
    Do While Forms.Count > 0
        DoCmd.Close acForm, Forms(0).Name
        DoEvents
    Loop
    Do While Reports.Count > 0
        DoCmd.Close acReport, Reports(0).Name
        DoEvents
    Loop
    Exit Function

errorHandler:
    Debug.Print "AppCodeImportExport.CloseFormsReports: Error #" & Err.Number & vbCrLf & Err.Description
End Function


'errno 457 - duplicate key (& item)
Public Function StrSetToCol(strSet As String, delimiter As String) As Collection 'throws errors
    Dim strSetArray() As String
    Dim Col As New Collection
    strSetArray = Split(strSet, delimiter)
    Dim Item As Variant
    For Each Item In strSetArray
        Col.Add Item, Item
    Next
    Set StrSetToCol = Col
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
' Author    : awaller
' Date      : 12/12/2016
' Purpose   : Save string variable to text file.
'---------------------------------------------------------------------------------------
'
Public Sub WriteFile(strContent As String, strPath As String)

    Dim fso As New Scripting.FileSystemObject
    Dim OutFile As Scripting.TextStream
    Set OutFile = fso.CreateTextFile(strPath, True)
    OutFile.Write strContent
    OutFile.Close
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetSafeFileName
' Author    : awaller
' Date      : 1/14/2019
' Purpose   : Replace illegal filename characters with URL encoded substitutes
'           : Sources: http://stackoverflow.com/questions/1976007/what-characters-are-forbidden-in-windows-and-linux-directory-names
'---------------------------------------------------------------------------------------
'
Public Function GetSafeFileName(strName As String) As String

    Dim strSafe As String
    
'    ' This created a lot of other issues when using the similar unicode characters.
'    strSafe = Replace(strName, "<", "«")
'    strSafe = Replace(strSafe, ">", "»")
'    strSafe = Replace(strSafe, ":", ChrW(760))
'    strSafe = Replace(strSafe, """", "”")
'    strSafe = Replace(strSafe, "/", ChrW(8725)) ' Division character
'    strSafe = Replace(strSafe, "\", ChrW(1633)) ' Arabic 1
'    strSafe = Replace(strSafe, "|", ChrW(1472)) ' Hebrew character
'    strSafe = Replace(strSafe, "?", ChrW(8253))
'    strSafe = Replace(strSafe, "*", ChrW(1645)) ' Arabic five-pointed star

    ' Instead, use URL encoding for these characters
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


Public Function StripDboPrefix(strName As String) As String
    If Left(strName, 4) = "dbo." Then
        StripDboPrefix = Mid(strName, 5)
    Else
        StripDboPrefix = strName
    End If
End Function