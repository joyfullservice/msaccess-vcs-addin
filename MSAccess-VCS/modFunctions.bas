Option Compare Database
Option Private Module
Option Explicit

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


' Can we export without closing the form?

' Export a database object with optional UCS2-to-UTF-8 conversion.
Public Sub ExportObject(obj_type_num As Integer, obj_name As String, file_path As String, _
    Optional Ucs2Convert As Boolean = False)

    modFunctions.VerifyPath Left(file_path, InStrRev(file_path, "\"))
    If Ucs2Convert Then
        Dim tempFileName As String: tempFileName = modFileAccess.TempFile()
        Application.SaveAsText obj_type_num, obj_name, tempFileName
        modFileAccess.ConvertUcs2Utf8 tempFileName, file_path
    Else
        Application.SaveAsText obj_type_num, obj_name, file_path
    End If
    If ShowDebugInfo Then Debug.Print "  " & obj_name
    
End Sub


' Import a database object with optional UTF-8-to-UCS2 conversion.
Public Sub ImportObject(obj_type_num As Integer, obj_name As String, file_path As String, _
    Optional Ucs2Convert As Boolean = False)
    
    If Not modFunctions.FileExists(file_path) Then Exit Sub
    
    If Ucs2Convert Then
        Dim tempFileName As String: tempFileName = modFileAccess.TempFile()
        modFileAccess.ConvertUtf8Ucs2 file_path, tempFileName
        Application.LoadFromText obj_type_num, obj_name, tempFileName
        
        Dim FSO As Object
        Set FSO = CreateObject("Scripting.FileSystemObject")
        FSO.DeleteFile tempFileName
    Else
        Application.LoadFromText obj_type_num, obj_name, file_path
    End If
End Sub

'shouldn't this be SanitizeTextFile (Singular)?

' For each *.txt in `Path`, find and remove a number of problematic but
' unnecessary lines of VB code that are inserted automatically by the
' Access GUI and change often (we don't want these lines of code in
' version control).
Public Sub SanitizeTextFiles(Path As String, Ext As String)


    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    '
    '  Setup Block matching Regex.
    Dim rxBlock As Object
    Set rxBlock = CreateObject("VBScript.RegExp")
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
    Dim rxLine As Object
    Set rxLine = CreateObject("VBScript.RegExp")
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
    Dim fileName As String
    fileName = Dir(Path & "*." & Ext)
    Dim isReport As Boolean: isReport = False
    Do Until Len(fileName) = 0
        DoEvents
        Dim obj_name As String
        obj_name = Mid(fileName, 1, InStrRev(fileName, ".") - 1)

        Dim InFile As Object
        Set InFile = FSO.OpenTextFile(Path & obj_name & "." & Ext, ForReading)
        Dim OutFile As Object
        Set OutFile = FSO.CreateTextFile(Path & obj_name & ".sanitize", True)
    
        Dim getLine As Boolean: getLine = True
        Do Until InFile.AtEndOfStream
            DoEvents
            Dim txt As String
            '
            ' Check if we need to get a new line of text
            If getLine = True Then
                txt = InFile.ReadLine
            Else
                getLine = True
            End If
            '
            ' Skip lines starting with line pattern
            If rxLine.Test(txt) Then
                Dim rxIndent As Object
                Set rxIndent = CreateObject("VBScript.RegExp")
                rxIndent.Pattern = "^(\s+)\S"
                '
                ' Get indentation level.
                Dim matches As Object
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
                    If rxIndent.Test(txt) Then Exit Do
                Loop
                ' We've moved on at least one line so do get a new one
                ' when starting the loop again.
                getLine = False
            '
            ' skip blocks of code matching block pattern
            ElseIf rxBlock.Test(txt) Then
                Do Until InFile.AtEndOfStream
                    txt = InFile.ReadLine
                    If InStr(txt, "End") Then Exit Do
                Loop
            ElseIf InStr(1, txt, "Begin Report") = 1 Then
                isReport = True
                OutFile.WriteLine txt
            ElseIf isReport = True And (InStr(1, txt, "    Right =") Or InStr(1, txt, "    Bottom =")) Then
                'skip line
                If InStr(1, txt, "    Bottom =") Then
                    isReport = False
                End If
            Else
                OutFile.WriteLine txt
            End If
        Loop
        OutFile.Close
        InFile.Close

        FSO.DeleteFile (Path & fileName)

        Dim thisFile As Object
        Set thisFile = FSO.GetFile(Path & obj_name & ".sanitize")
        thisFile.Move (Path & fileName)
        fileName = Dir()
    Loop


End Sub




' Path/Directory of the current database file.
Public Function ProjectPath() As String
    ProjectPath = CurrentProject.Path
    If Right(ProjectPath, 1) <> "\" Then ProjectPath = ProjectPath & "\"
End Function

' Path/Directory for source files
Public Function SourcePath() As String
    SourcePath = ProjectPath & CurrentProject.name & ".src\"
End Function

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
Public Sub ClearTextFilesFromDir(Path As String, Ext As String)
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If Not FSO.FolderExists(Path) Then Exit Sub

    On Error GoTo ClearTextFilesFromDir_noop
    If Dir(Path & "*." & Ext) <> "" Then
        FSO.DeleteFile Path & "*." & Ext
    End If
ClearTextFilesFromDir_noop:

    On Error GoTo 0
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
    Dim x(-1 To -1) As String
    Sb_Init = x
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
Public Function SubString(p As Integer, s As String, startsWith As String, endsWith As String)
    Dim start As Integer
    Dim cursor As Integer
    Dim p1 As Integer
    Dim p2 As Integer
    Dim level As Integer
    start = InStr(p, s, startsWith)
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
        DoCmd.Close acForm, Forms(0).name
        DoEvents
    Loop
    Do While Reports.Count > 0
        DoCmd.Close acReport, Reports(0).name
        DoEvents
    Loop
    Exit Function

errorHandler:
    Debug.Print "AppCodeImportExport.CloseFormsReports: Error #" & Err.Number & vbCrLf & Err.Description
End Function


'errno 457 - duplicate key (& item)
Public Function StrSetToCol(strSet As String, delimiter As String) As Collection 'throws errors
    Dim strSetArray() As String
    Dim col As New Collection
    strSetArray = Split(strSet, delimiter)
    Dim item As Variant
    For Each item In strSetArray
        col.Add item, item
    Next
    Set StrSetToCol = col
End Function