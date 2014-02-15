Attribute VB_Name = "AppCodeImportExport"
' Access Module `AppCodeImportExport`
' -----------------------------------
'
' Version 0.9
'
' https://github.com/bkidwell/msaccess-vcs-integration
'
' Brendan Kidwell
'
' This code is licensed under BSD-style terms.
'
' This is some code for importing and exporting Access Queries, Forms,
' Reports, Macros, and Modules to and from plain text files, for the
' purpose of syncing with a version control system.
'
'
' Use:
'
' BACKUP YOUR WORK BEFORE TRYING THIS CODE!
'
' To create and/or overwrite source text files for all database objects
' (except tables) in "$database-folder/source/", run
' `ExportAllSource()`.
'
' To load and/or overwrite  all database objects from source files in
' "$database-folder/source/", run `ImportAllSource()`.
'
' See project home page (URL above) for more information.
'
'
' Future expansion:
' * Maybe integrate into a dialog box triggered by a menu item.
' * Warning of destructive overwrite.

Option Compare Database
Option Explicit

' --------------------------------
' Configuration
' --------------------------------

' List of lookup tables that are part of the program rather than the
' data, to be exported with source code
'
' Provide a comma separated list of table names, or an empty string
' ("") if no tables are to be exported with the source code.

Private Const INCLUDE_TABLES = ""

' Do more aggressive removal of superfluous blobs from exported MS
' Access source code?

Private Const AggressiveSanitize = True
Private Const StripPublishOption = True
'
' --------------------------------
' Structures
' --------------------------------

' Structure to track buffered reading or writing of binary files
Private Type BinFile
    file_num As Integer
    file_len As Long
    file_pos As Long
    buffer As String
    buffer_len As Integer
    buffer_pos As Integer
    at_eof As Boolean
    mode As String
End Type

' --------------------------------
' Constants
' --------------------------------

' Constants for Scripting.FileSystemObject API
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const TristateTrue = -1, TristateFalse = 0, TristateUseDefault = -2

' --------------------------------
' Module variables
' --------------------------------
'
' Does the current database file write UCS2-little-endian when exporting
' Queries, Forms, Reports, Macros
Private UsingUcs2 As Boolean
'
' --------------------------------
' External Library Functions
' --------------------------------
Private Declare PtrSafe _
    Function getTempPath Lib "kernel32" _
         Alias "GetTempPathA" (ByVal nBufferLength As Long, _
                               ByVal lpBuffer As String) As Long
Private Declare PtrSafe _
    Function getTempFileName Lib "kernel32" _
         Alias "GetTempFileNameA" (ByVal lpszPath As String, _
                                   ByVal lpPrefixString As String, _
                                   ByVal wUnique As Long, _
                                   ByVal lpTempFileName As String) As Long
' --------------------------------
' Basic functions missing from VB 6: buffered file read/write, string builder
' --------------------------------

' Open a binary file for reading (mode = 'r') or writing (mode = 'w').
Private Function BinOpen(file_path As String, mode As String) As BinFile
    Dim F As BinFile

    F.file_num = FreeFile
    F.mode = LCase(mode)
    If F.mode = "r" Then
        Open file_path For Binary Access Read As F.file_num
        F.file_len = LOF(F.file_num)
        F.file_pos = 0
        If F.file_len > &H4000 Then
            F.buffer = String(&H4000, " ")
            F.buffer_len = &H4000
        Else
            F.buffer = String(F.file_len, " ")
            F.buffer_len = F.file_len
        End If
        F.buffer_pos = 0
        Get F.file_num, F.file_pos + 1, F.buffer
    Else
        DelIfExist file_path
        Open file_path For Binary Access Write As F.file_num
        F.file_len = 0
        F.file_pos = 0
        F.buffer = String(&H4000, " ")
        F.buffer_len = 0
        F.buffer_pos = 0
    End If

    BinOpen = F
End Function

' Buffered read one byte at a time from a binary file.
Private Function BinRead(ByRef F As BinFile) As Integer
    If F.at_eof = True Then
        BinRead = 0
        Exit Function
    End If

    BinRead = Asc(Mid(F.buffer, F.buffer_pos + 1, 1))

    F.buffer_pos = F.buffer_pos + 1
    If F.buffer_pos >= F.buffer_len Then
        F.file_pos = F.file_pos + &H4000
        If F.file_pos >= F.file_len Then
            F.at_eof = True
            Exit Function
        End If
        If F.file_len - F.file_pos > &H4000 Then
            F.buffer_len = &H4000
        Else
            F.buffer_len = F.file_len - F.file_pos
            F.buffer = String(F.buffer_len, " ")
        End If
        F.buffer_pos = 0
        Get F.file_num, F.file_pos + 1, F.buffer
    End If
End Function

' Buffered write one byte at a time from a binary file.
Private Sub BinWrite(ByRef F As BinFile, b As Integer)
    Mid(F.buffer, F.buffer_pos + 1, 1) = Chr(b)
    F.buffer_pos = F.buffer_pos + 1
    If F.buffer_pos >= &H4000 Then
        Put F.file_num, , F.buffer
        F.buffer_pos = 0
    End If
End Sub

' Close binary file.
Private Sub BinClose(ByRef F As BinFile)
    If F.mode = "w" And F.buffer_pos > 0 Then
        F.buffer = Left(F.buffer, F.buffer_pos)
        Put F.file_num, , F.buffer
    End If
    Close F.file_num
End Sub

' String builder: Init
Private Function Sb_Init() As String()
    Dim x(-1 To -1) As String
    Sb_Init = x
End Function

' String builder: Clear
Private Sub Sb_Clear(ByRef sb() As String)
    ReDim Sb_Init(-1 To -1)
End Sub

' String builder: Append
Private Sub Sb_Append(ByRef sb() As String, Value As String)
    If LBound(sb) = -1 Then
        ReDim sb(0 To 0)
    Else
        ReDim Preserve sb(0 To UBound(sb) + 1)
    End If
    sb(UBound(sb)) = Value
End Sub

' String builder: Get value
Private Function Sb_Get(ByRef sb() As String) As String
    Sb_Get = Join(sb, "")
End Function

' --------------------------------
' Beginning of main functions of this module
' --------------------------------

' Close all open forms.
Private Function CloseFormsReports()
    On Error GoTo errorHandler
    Do While Forms.count > 0
        DoCmd.Close acForm, Forms(0).Name
    Loop
    Do While Reports.count > 0
        DoCmd.Close acReport, Reports(0).Name
    Loop
    Exit Function

errorHandler:
    Debug.Print "AppCodeImportExport.CloseFormsReports: Error #" & Err.Number & vbCrLf & Err.description
End Function

' Pad a string on the right to make it `count` characters long.
Public Function PadRight(Value As String, count As Integer)
    PadRight = Value
    If Len(Value) < count Then
        PadRight = PadRight & Space(count - Len(Value))
    End If
End Function

' Path of the current database file.
Private Function ProjectPath() As String
    ProjectPath = CurrentProject.Path
    If Right(ProjectPath, 1) <> "\" Then ProjectPath = ProjectPath & "\"
End Function
'
' Generate Random / Unique tempprary file name.
Private Function TempFile(Optional sPrefix As String = "VBA") As String
Dim sTmpPath As String * 512
Dim sTmpName As String * 576
Dim nRet As Long
Dim sFileName As String
    
    nRet = getTempPath(512, sTmpPath)
    nRet = getTempFileName(sTmpPath, sPrefix, 0, sTmpName)
    If nRet <> 0 Then sFileName = Left$(sTmpName, InStr(sTmpName, vbNullChar) - 1)
    TempFile = sFileName
End Function

' Export a database object with optional UCS2-to-UTF-8 conversion.
Private Sub ExportObject(obj_type_num As Integer, obj_name As String, file_path As String, _
    Optional Ucs2Convert As Boolean = False)

    MkDirIfNotExist Left(file_path, InStrRev(file_path, "\"))
    If Ucs2Convert Then
        Dim tempFileName As String: tempFileName = TempFile()
        Application.SaveAsText obj_type_num, obj_name, tempFileName
        ConvertUcs2Utf8 tempFileName, file_path
    Else
        Application.SaveAsText obj_type_num, obj_name, file_path
    End If
End Sub

' Import a database object with optional UTF-8-to-UCS2 conversion.
Public Sub ImportObject(obj_type_num As Integer, obj_name As String, file_path As String, _
    Optional Ucs2Convert As Boolean = False)

    If Ucs2Convert Then
        Dim tempFileName As String: tempFileName = TempFile()
        ConvertUtf8Ucs2 file_path, tempFileName
        Application.LoadFromText obj_type_num, obj_name, tempFileName
        DeleteFile tempFileName
    Else
        Application.LoadFromText obj_type_num, obj_name, file_path
    End If
End Sub

' Binary convert a UCS2-little-endian encoded file to UTF-8.
Private Sub ConvertUcs2Utf8(Source As String, dest As String)
    Dim f_in As BinFile, f_out As BinFile
    Dim in_low As Integer, in_high As Integer

    f_in = BinOpen(Source, "r")
    f_out = BinOpen(dest, "w")

    Do While Not f_in.at_eof
        in_low = BinRead(f_in)
        in_high = BinRead(f_in)
        If in_high = 0 And in_low < &H80 Then
            ' U+0000 - U+007F   0LLLLLLL
            BinWrite f_out, in_low
        ElseIf in_high < &H8 Then
            ' U+0080 - U+07FF   110HHHLL 10LLLLLL
            BinWrite f_out, &HC0 + ((in_high And &H7) * &H4) + ((in_low And &HC0) / &H40)
            BinWrite f_out, &H80 + (in_low And &H3F)
        Else
            ' U+0800 - U+FFFF   1110HHHH 10HHHHLL 10LLLLLL
            BinWrite f_out, &HE0 + ((in_high And &HF0) / &H10)
            BinWrite f_out, &H80 + ((in_high And &HF) * &H4) + ((in_low And &HC0) / &H40)
            BinWrite f_out, &H80 + (in_low And &H3F)
        End If
    Loop

    BinClose f_in
    BinClose f_out
End Sub

' Binary convert a UTF-8 encoded file to UCS2-little-endian.
Private Sub ConvertUtf8Ucs2(Source As String, dest As String)
    Dim f_in As BinFile, f_out As BinFile
    Dim in_1 As Integer, in_2 As Integer, in_3 As Integer

    f_in = BinOpen(Source, "r")
    f_out = BinOpen(dest, "w")

    Do While Not f_in.at_eof
        in_1 = BinRead(f_in)
        If (in_1 And &H80) = 0 Then
            ' U+0000 - U+007F   0LLLLLLL
            BinWrite f_out, in_1
            BinWrite f_out, 0
        ElseIf (in_1 And &HE0) = &HC0 Then
            ' U+0080 - U+07FF   110HHHLL 10LLLLLL
            in_2 = BinRead(f_in)
            BinWrite f_out, ((in_1 And &H3) * &H40) + (in_2 And &H3F)
            BinWrite f_out, (in_1 And &H1C) / &H4
        Else
            ' U+0800 - U+FFFF   1110HHHH 10HHHHLL 10LLLLLL
            in_2 = BinRead(f_in)
            in_3 = BinRead(f_in)
            BinWrite f_out, ((in_2 And &H3) * &H40) + (in_3 And &H3F)
            BinWrite f_out, ((in_1 And &HF) * &H10) + ((in_2 And &H3C) / &H4)
        End If
    Loop

    BinClose f_in
    BinClose f_out
End Sub

' Determine if this database imports/exports code as UCS-2-LE. (Older file
' formats cause exported objects to use a Windows 8-bit character set.)
Private Sub InitUsingUcs2()
    Dim obj_name As String, i As Integer, obj_type As Variant, fn As Integer, bytes As String
    Dim obj_type_split() As String, obj_type_name As String, obj_type_num As Integer
    Dim db As Object ' DAO.Database

    If CurrentDb.QueryDefs.count > 0 Then
        obj_type_num = acQuery
        obj_name = CurrentDb.QueryDefs(0).Name
    Else
        For Each obj_type In Split( _
            "Forms|" & acForm & "," & _
            "Reports|" & acReport & "," & _
            "Scripts|" & acMacro & "," & _
            "Modules|" & acModule _
        )
            obj_type_split = Split(obj_type, "|")
            obj_type_name = obj_type_split(0)
            obj_type_num = Val(obj_type_split(1))
            If CurrentDb.Containers(obj_type_name).Documents.count > 0 Then
                obj_name = CurrentDb.Containers(obj_type_name).Documents(1).Name
                Exit For
            End If
        Next
    End If

    If obj_name = "" Then
        ' No objects found that can be used to test UCS2 versus UTF-8
        UsingUcs2 = True
        Exit Sub
    End If

    Dim tempFileName As String: tempFileName = TempFile()
    Application.SaveAsText obj_type_num, obj_name, tempFileName
    fn = FreeFile
    Open tempFileName For Binary Access Read As fn
    bytes = "  "
    Get fn, 1, bytes
    If Asc(Mid(bytes, 1, 1)) = &HFF And Asc(Mid(bytes, 2, 1)) = &HFE Then
        UsingUcs2 = True
    Else
        UsingUcs2 = False
    End If
    Close fn
    DeleteFile (tempFileName)
End Sub

' Create folder `Path`. Silently do nothing if it already exists.
Private Sub MkDirIfNotExist(Path As String)
    On Error GoTo MkDirIfNotexist_noop
    MkDir Path
MkDirIfNotexist_noop:
    On Error GoTo 0
End Sub

' Delete a file if it exists.
Private Sub DelIfExist(Path As String)
    On Error GoTo DelIfNotExist_Noop
    Kill Path
DelIfNotExist_Noop:
    On Error GoTo 0
End Sub

' Erase all *.`ext` files in `Path`.
Private Sub ClearTextFilesFromDir(Path As String, Ext As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(Path) Then Exit Sub

    On Error GoTo ClearTextFilesFromDir_noop
    If Dir(Path & "*." & Ext) <> "" Then
        Kill Path & "*." & Ext
    End If
ClearTextFilesFromDir_noop:

    On Error GoTo 0
End Sub

' For each *.txt in `Path`, find and remove a number of problematic but
' unnecessary lines of VB code that are inserted automatically by the
' Access GUI and change often (we don't want these lines of code in
' version control).
Private Sub SanitizeTextFiles(Path As String, Ext As String)
Dim fso As Object
Dim thisFile As Object
Dim InFile As Object
Dim OutFile As Object
Dim fileName As String
Dim txt As String
Dim obj_name As String
Dim rxBlock As Object
Dim rxLine As Object
Dim matches As String
Dim deleteCount As Integer

    Set fso = CreateObject("Scripting.FileSystemObject")
    '
    '  Setup Block matching Regex.
    Set rxBlock = CreateObject("VBScript.RegExp")
    rxBlock.ignoreCase = False
    '
    '  Match PrtDevNames / Mode with or  without W
    matches = "PrtDev(?:Names|Mode)[W]?"
    If (AggressiveSanitize = True) Then
      '  Add and group aggressive matches
      matches = "(?:" & matches
      matches = matches & "|GUID|NameMap|dbLongBinary ""DOL"""
      matches = matches & ")"
    End If
    '  Ensure that this is the begining of a block.
    matches = matches & " = Begin"
'Debug.Print matches
    rxBlock.pattern = matches
    '
    '  Setup Line Matching Regex.
    Set rxLine = CreateObject("VBScript.RegExp")
    matches = "^\s*(?:"
    matches = matches & "Checksum ="
    matches = matches & "|BaseInfo|NoSaveCTIWhenDisabled =1"
    If (StripPublishOption = True) Then
        matches = matches & "|dbByte ""PublishToWeb"" =""1"""
        matches = matches & "|PublishOption =1"
    End If
    matches = matches & ")"
'Debug.Print matches
    rxLine.pattern = matches
    
    fileName = Dir(Path & "*." & Ext)
    Do Until Len(fileName) = 0
        obj_name = Mid(fileName, 1, InStrRev(fileName, ".") - 1)

        Set InFile = fso.OpenTextFile(Path & obj_name & "." & Ext, ForReading)
        Set OutFile = fso.CreateTextFile(Path & obj_name & ".sanitize", True)
        Do Until InFile.AtEndOfStream
            txt = InFile.ReadLine
                '
                ' Skip lines starting with line pattern
            If rxLine.Test(txt) Then
                '
                ' skip blocks of code matching block pattern
            ElseIf rxBlock.Test(txt) Then
                Do Until InFile.AtEndOfStream
                    txt = InFile.ReadLine
                    If InStr(txt, "End") Then Exit Do
                Loop
            Else
                OutFile.WriteLine txt
            End If
        Loop
        OutFile.Close
        InFile.Close

        DeleteFile (Path & fileName)

        Set thisFile = fso.GetFile(Path & obj_name & ".sanitize")
        thisFile.Move (Path & fileName)
        fileName = Dir()
    Loop

End Sub

' Main entry point for EXPORT. Export all forms, reports, queries,
' macros, modules, and lookup tables to `source` folder under the
' database's folder.
Public Sub ExportAllSource()
    Dim db As Object ' DAO.Database
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
    Dim ucs2 As Boolean
    Dim tblName As Variant

    Set db = CurrentDb

    CloseFormsReports
    InitUsingUcs2

    source_path = ProjectPath() & "source\"
    MkDirIfNotExist source_path

    Debug.Print

    obj_path = source_path & "queries\"
    ClearTextFilesFromDir obj_path, "bas"
    Debug.Print PadRight("Exporting queries...", 24);
    obj_count = 0
    For Each qry In db.QueryDefs
        If Left(qry.Name, 1) <> "~" Then
            ExportObject acQuery, qry.Name, obj_path & qry.Name & ".bas", UsingUcs2
            obj_count = obj_count + 1
        End If
    Next
    SanitizeTextFiles obj_path, "bas"
    Debug.Print "[" & obj_count & "]"

    obj_path = source_path & "tables\"
    ClearTextFilesFromDir obj_path, "txt"
    If (Len(Replace(INCLUDE_TABLES, " ", "")) > 0) Then
        Debug.Print PadRight("Exporting tables...", 24);
        obj_count = 0
        For Each tblName In Split(INCLUDE_TABLES, ",")
            ExportTable CStr(tblName), obj_path
            obj_count = obj_count + 1
        Next
        Debug.Print "[" & obj_count & "]"
    End If

    For Each obj_type In Split( _
        "forms|Forms|" & acForm & "," & _
        "reports|Reports|" & acReport & "," & _
        "macros|Scripts|" & acMacro & "," & _
        "modules|Modules|" & acModule _
        , "," _
    )
        obj_type_split = Split(obj_type, "|")
        obj_type_label = obj_type_split(0)
        obj_type_name = obj_type_split(1)
        obj_type_num = Val(obj_type_split(2))
        obj_path = source_path & obj_type_label & "\"
        obj_count = 0
        ClearTextFilesFromDir obj_path, "bas"
        Debug.Print PadRight("Exporting " & obj_type_label & "...", 24);
        For Each doc In db.Containers(obj_type_name).Documents
            If (Left(doc.Name, 1) <> "~") And _
               (doc.Name <> "AppCodeImportExport") Then
                If obj_type_label = "modules" Then
                    ucs2 = False
                Else
                    ucs2 = UsingUcs2
                End If
                ExportObject obj_type_num, doc.Name, obj_path & doc.Name & ".bas", ucs2
                obj_count = obj_count + 1
            End If
        Next
        Debug.Print "[" & obj_count & "]"

        If obj_type_label <> "modules" Then
            SanitizeTextFiles obj_path, "bas"
        End If
    Next

    Debug.Print "Done."
End Sub

' Main entry point for IMPORT. Import all forms, reports, queries,
' macros, modules, and lookup tables from `source` folder under the
' database's folder.
Public Sub ImportAllSource()
    Dim db As Object ' DAO.Database
    Dim fso As Object
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
    Dim fileName As String
    Dim obj_name As String
    Dim ucs2 As Boolean

    Set db = CurrentDb
    Set fso = CreateObject("Scripting.FileSystemObject")

    CloseFormsReports
    InitUsingUcs2

    source_path = ProjectPath() & "source\"
    If Not fso.FolderExists(source_path) Then
        MsgBox "No source found at:" & vbCrLf & source_path, vbExclamation, "Import failed"
        Exit Sub
    End If

    Debug.Print

    obj_path = source_path & "queries\"
    fileName = Dir(obj_path & "*.bas")
    If Len(fileName) > 0 Then
        Debug.Print PadRight("Importing queries...", 24);
        obj_count = 0
        Do Until Len(fileName) = 0
            obj_name = Mid(fileName, 1, InStrRev(fileName, ".") - 1)
            ImportObject acQuery, obj_name, obj_path & fileName, UsingUcs2
            obj_count = obj_count + 1
            fileName = Dir()
        Loop
        Debug.Print "[" & obj_count & "]"
    End If

    obj_path = source_path & "tables\"
    fileName = Dir(obj_path & "*.txt")
    If Len(fileName) > 0 Then
        Debug.Print PadRight("Importing tables...", 24);
        obj_count = 0
        Do Until Len(fileName) = 0
            obj_name = Mid(fileName, 1, InStrRev(fileName, ".") - 1)
            ImportTable CStr(obj_name), obj_path
            obj_count = obj_count + 1
            fileName = Dir()
        Loop
        Debug.Print "[" & obj_count & "]"
    End If

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
            Debug.Print PadRight("Importing " & obj_type_label & "...", 24);
            obj_count = 0
            Do Until Len(fileName) = 0
                obj_name = Mid(fileName, 1, InStrRev(fileName, ".") - 1)
                If obj_name <> "AppCodeImportExport" Then
                    If obj_type_label = "modules" Then
                        ucs2 = False
                    Else
                        ucs2 = UsingUcs2
                    End If
                    ImportObject obj_type_num, obj_name, obj_path & fileName, ucs2
                    obj_count = obj_count + 1
                End If
                fileName = Dir()
            Loop
            Debug.Print "[" & obj_count & "]"
        End If
    Next

    Debug.Print "Done."
End Sub

' Build SQL to export `tbl_name` sorted by each field from first to last
Public Function TableExportSql(tbl_name As String)
    Dim rs As Object ' DAO.Recordset
    Dim fieldObj As Object ' DAO.Field
    Dim sb() As String, count As Integer

    Set rs = CurrentDb.OpenRecordset(tbl_name)

    sb = Sb_Init()
    Sb_Append sb, "SELECT "
    count = 0
    For Each fieldObj In rs.Fields
        If count > 0 Then Sb_Append sb, ", "
        Sb_Append sb, "[" & fieldObj.Name & "]"
        count = count + 1
    Next
    Sb_Append sb, " FROM [" & tbl_name & "] ORDER BY "
    count = 0
    For Each fieldObj In rs.Fields
        If count > 0 Then Sb_Append sb, ", "
        Sb_Append sb, "[" & fieldObj.Name & "]"
        count = count + 1
    Next

    TableExportSql = Sb_Get(sb)
End Function

' Export the lookup table `tblName` to `source\tables`.
Private Sub ExportTable(tbl_name As String, obj_path As String)
    Dim fso, OutFile
    Dim rs As Object ' DAO.Recordset
    Dim fieldObj As Object ' DAO.Field
    Dim C As Long, Value As Variant

    Set fso = CreateObject("Scripting.FileSystemObject")
    ' open file for writing with Create=True, Unicode=True (USC-2 Little Endian format)
    MkDirIfNotExist obj_path
    Dim tempFileName As String: tempFileName = TempFile()

    Set OutFile = fso.CreateTextFile(tempFileName, True, True)

    Set rs = CurrentDb.OpenRecordset(TableExportSql(tbl_name))
    C = 0
    For Each fieldObj In rs.Fields
        If C <> 0 Then OutFile.Write vbTab
        C = C + 1
        OutFile.Write fieldObj.Name
    Next
    OutFile.Write vbCrLf

    rs.MoveFirst
    Do Until rs.EOF
        C = 0
        For Each fieldObj In rs.Fields
            If C <> 0 Then OutFile.Write vbTab
            C = C + 1
            Value = rs(fieldObj.Name)
            If IsNull(Value) Then
                Value = ""
            Else
                Value = Replace(Value, "\", "\\")
                Value = Replace(Value, vbCrLf, "\n")
                Value = Replace(Value, vbCr, "\n")
                Value = Replace(Value, vbLf, "\n")
                Value = Replace(Value, vbTab, "\t")
            End If
            OutFile.Write Value
        Next
        OutFile.Write vbCrLf
        rs.MoveNext
    Loop
    rs.Close
    OutFile.Close

    ConvertUcs2Utf8 tempFileName, obj_path & tbl_name & ".txt"
    DeleteFile tempFileName
End Sub

' Import the lookup table `tblName` from `source\tables`.
Private Sub ImportTable(tblName As String, obj_path As String)
    Dim db As Object ' DAO.Database
    Dim rs As Object ' DAO.Recordset
    Dim fieldObj As Object ' DAO.Field
    Dim fso, InFile As Object
    Dim C As Long, buf As String, Values() As String, Value As Variant

    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim tempFileName As String: tempFileName = TempFile()
    ConvertUtf8Ucs2 obj_path & tblName & ".txt", tempFileName
    ' open file for reading with Create=False, Unicode=True (USC-2 Little Endian format)
    Set InFile = fso.OpenTextFile(tempFileName, ForReading, False, TristateTrue)
    Set db = CurrentDb

    db.Execute "DELETE FROM [" & tblName & "]"
    Set rs = db.OpenRecordset(tblName)
    buf = InFile.ReadLine()
    Do Until InFile.AtEndOfStream
        buf = InFile.ReadLine()
        If Len(Trim(buf)) > 0 Then
            Values = Split(buf, vbTab)
            C = 0
            rs.AddNew
            For Each fieldObj In rs.Fields
                Value = Values(C)
                If Len(Value) = 0 Then
                    Value = Null
                Else
                    Value = Replace(Value, "\t", vbTab)
                    Value = Replace(Value, "\n", vbCrLf)
                    Value = Replace(Value, "\\", "\")
                End If
                rs(fieldObj.Name) = Value
                C = C + 1
            Next
            rs.Update
        End If
    Loop

    rs.Close
    InFile.Close
    DeleteFile tempFileName
End Sub

Private Function DeleteFile(sFileName As String)
Dim deleteCount As Integer
Dim pauseStart As Variant
Dim pauseEnd As Variant
'
'  Try to delete the file a few times if it fails.
'  Failure is generally caused by the file not actually being closed yet.
        deleteCount = 0
        On Error GoTo tryDeleteAgain
tryDeleteAgain:
        If deleteCount > 0 Then
            pauseStart = Timer
            pauseEnd = pauseStart + 0.1
            Do While Timer < pauseEnd
                DoEvents
            Loop
        ElseIf deleteCount > 3 Then
            On Error GoTo 0
        End If
        deleteCount = deleteCount + 1
        Kill sFileName
'        If deleteCount > 1 Then Debug.Print "Delete Attempts [" & deleteCount & "] (" & sFileName & ")"
        '
        '  Release Error Handler
        On Error GoTo 0

End Function
'
' Main entry point for ImportProject.
' Drop all forms, reports, queries, macros, modules.
' execute ImportAllSource.
Public Sub ImportProject()
On Error GoTo errorHandler

    If MsgBox("This action will delete all existing: " & vbCrLf & _
              vbCrLf & _
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

    Dim db As DAO.Database
    Set db = CurrentDb
    CloseFormsReports

    Debug.Print
    Debug.Print "Deleting Existing Objects"
    Debug.Print

    Dim dbObject As Object
    For Each dbObject In db.QueryDefs
        If Left(dbObject.Name, 1) <> "~" Then
'            Debug.Print dbObject.Name
            db.QueryDefs.Delete dbObject.Name
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

        For Each doc In db.Containers(objTypeArray(OTNAME)).Documents
            If (Left(doc.Name, 1) <> "~") And _
               (doc.Name <> "AppCodeImportExport") Then
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
  Debug.Print "AppCodeImportExport.ImportProject: Error #" & Err.Number & vbCrLf & _
               Err.description

exitHandler:
End Sub
