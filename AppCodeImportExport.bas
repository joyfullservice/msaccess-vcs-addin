' Access Module `AppCodeImportExport`
' -----------------------------------
'
' Version 0.3
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
' Structure to keep track of "on Update" and "on Delete" clauses
' Access does not in all cases execute such queries
Private Type structEnforce
    foreignTable As String
    foreignFields() As String
    Table As String
    refFields() As String
    isUpdate As Boolean
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

' Does the current database file write UCS2-little-endian when exporting
' Queries, Forms, Reports, Macros
Private UsingUcs2 As Boolean

' keeping "on Update" relations to be complemented afte table creation
Private K() As structEnforce


' --------------------------------
' Basic functions missing from VB 6: buffered file read/write, string builder
' --------------------------------

' Open a binary file for reading (mode = 'r') or writing (mode = 'w').
Private Function BinOpen(file_path As String, mode As String) As BinFile
    Dim f As BinFile

    f.file_num = FreeFile
    f.mode = LCase(mode)
    If f.mode = "r" Then
        Open file_path For Binary Access Read As f.file_num
        f.file_len = LOF(f.file_num)
        f.file_pos = 0
        If f.file_len > &H4000 Then
            f.buffer = String(&H4000, " ")
            f.buffer_len = &H4000
        Else
            f.buffer = String(f.file_len, " ")
            f.buffer_len = f.file_len
        End If
        f.buffer_pos = 0
        Get f.file_num, f.file_pos + 1, f.buffer
    Else
        DelIfExist file_path
        Open file_path For Binary Access Write As f.file_num
        f.file_len = 0
        f.file_pos = 0
        f.buffer = String(&H4000, " ")
        f.buffer_len = 0
        f.buffer_pos = 0
    End If

    BinOpen = f
End Function

' Buffered read one byte at a time from a binary file.
Private Function BinRead(ByRef f As BinFile) As Integer
    If f.at_eof = True Then
        BinRead = 0
        Exit Function
    End If

    BinRead = Asc(Mid(f.buffer, f.buffer_pos + 1, 1))

    f.buffer_pos = f.buffer_pos + 1
    If f.buffer_pos >= f.buffer_len Then
        f.file_pos = f.file_pos + &H4000
        If f.file_pos >= f.file_len Then
            f.at_eof = True
            Exit Function
        End If
        If f.file_len - f.file_pos > &H4000 Then
            f.buffer_len = &H4000
        Else
            f.buffer_len = f.file_len - f.file_pos
            f.buffer = String(f.buffer_len, " ")
        End If
        f.buffer_pos = 0
        Get f.file_num, f.file_pos + 1, f.buffer
    End If
End Function

' Buffered write one byte at a time from a binary file.
Private Sub BinWrite(ByRef f As BinFile, b As Integer)
    Mid(f.buffer, f.buffer_pos + 1, 1) = Chr(b)
    f.buffer_pos = f.buffer_pos + 1
    If f.buffer_pos >= &H4000 Then
        Put f.file_num, , f.buffer
        f.buffer_pos = 0
    End If
End Sub

' Close binary file.
Private Sub BinClose(ByRef f As BinFile)
    If f.mode = "w" And f.buffer_pos > 0 Then
        f.buffer = Left(f.buffer, f.buffer_pos)
        Put f.file_num, , f.buffer
    End If
    Close f.file_num
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
    On Error GoTo ErrorHandler
    Do While Forms.count > 0
        DoCmd.Close acForm, Forms(0).name
    Loop
    Do While Reports.count > 0
        DoCmd.Close acReport, Reports(0).name
    Loop
    Exit Function

ErrorHandler:
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

' Path of single temp file used by any function in this module.
Private Function TempFile() As String
    TempFile = ProjectPath() & "AppCodeImportExport.tempdata"
End Function

' Export a database object with optional UCS2-to-UTF-8 conversion.
Private Sub ExportObject(obj_type_num As Integer, obj_name As String, file_path As String, _
    Optional Ucs2Convert As Boolean = False)

    MkDirIfNotExist Left(file_path, InStrRev(file_path, "\"))
    If Ucs2Convert Then
        Application.SaveAsText obj_type_num, obj_name, TempFile()
        ConvertUcs2Utf8 TempFile(), file_path
    Else
        Application.SaveAsText obj_type_num, obj_name, file_path
    End If
End Sub

' Import a database object with optional UTF-8-to-UCS2 conversion.
Private Sub ImportObject(obj_type_num As Integer, obj_name As String, file_path As String, _
    Optional Ucs2Convert As Boolean = False)

    If Ucs2Convert Then
        ConvertUtf8Ucs2 file_path, TempFile()
        Application.LoadFromText obj_type_num, obj_name, TempFile()
    Else
        Application.LoadFromText obj_type_num, obj_name, file_path
    End If
End Sub

' Binary convert a UCS2-little-endian encoded file to UTF-8.
Private Sub ConvertUcs2Utf8(source As String, dest As String)
    Dim f_in As BinFile, f_out As BinFile
    Dim in_low As Integer, in_high As Integer

    f_in = BinOpen(source, "r")
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
Private Sub ConvertUtf8Ucs2(source As String, dest As String)
    Dim f_in As BinFile, f_out As BinFile
    Dim in_1 As Integer, in_2 As Integer, in_3 As Integer

    f_in = BinOpen(source, "r")
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
    Dim Db As Object ' DAO.Database

    If CurrentDb.QueryDefs.count > 0 Then
        obj_type_num = acQuery
        obj_name = CurrentDb.QueryDefs(0).name
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
                obj_name = CurrentDb.Containers(obj_type_name).Documents(1).name
                Exit For
            End If
        Next
    End If

    If obj_name = "" Then
        ' No objects found that can be used to test UCS2 versus UTF-8
        UsingUcs2 = True
        Exit Sub
    End If

    Application.SaveAsText obj_type_num, obj_name, TempFile()
    fn = FreeFile
    Open TempFile() For Binary Access Read As fn
    bytes = "  "
    Get fn, 1, bytes
    If Asc(Mid(bytes, 1, 1)) = &HFF And Asc(Mid(bytes, 2, 1)) = &HFE Then
        UsingUcs2 = True
    Else
        UsingUcs2 = False
    End If
    Close fn
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
    Dim fso, InFile, OutFile, FileName As String, txt As String, obj_name As String

    Set fso = CreateObject("Scripting.FileSystemObject")

    FileName = Dir(Path & "*." & Ext)
    Do Until Len(FileName) = 0
        obj_name = Mid(FileName, 1, InStrRev(FileName, ".") - 1)

        Set InFile = fso.OpenTextFile(Path & obj_name & "." & Ext, ForReading)
        Set OutFile = fso.CreateTextFile(Path & obj_name & ".sanitize", True)
        Do Until InFile.AtEndOfStream
            txt = InFile.ReadLine
            If Left(txt, 10) = "Checksum =" Then
                ' Skip lines starting with Checksum
            ElseIf InStr(txt, "NoSaveCTIWhenDisabled =1") Then
                ' Skip lines containning NoSaveCTIWhenDisabled
            ElseIf InStr(txt, "Begin") > 0 Then
                If _
                    InStr(txt, "PrtDevNames =") > 0 Or _
                    InStr(txt, "PrtDevNamesW =") > 0 Or _
                    InStr(txt, "PrtDevModeW =") > 0 Or _
                    InStr(txt, "PrtDevMode =") > 0 _
                    Then

                    ' skip this block of code
                    Do Until InFile.AtEndOfStream
                        txt = InFile.ReadLine
                        If InStr(txt, "End") Then Exit Do
                    Loop
                ElseIf AggressiveSanitize And ( _
                    InStr(txt, "dbLongBinary ""DOL"" =") > 0 Or _
                    InStr(txt, "NameMap") > 0 Or _
                    InStr(txt, "GUID") > 0 _
                    ) Then

                    ' skip this block of code
                    Do Until InFile.AtEndOfStream
                        txt = InFile.ReadLine
                        If InStr(txt, "End") Then Exit Do
                    Loop
                Else
                    ' Something else has begun
                     OutFile.WriteLine txt
                End If
            Else
                OutFile.WriteLine txt
            End If
        Loop
        OutFile.Close
        InFile.Close

        FileName = Dir()
    Loop

    FileName = Dir(Path & "*." & Ext)
    Do Until Len(FileName) = 0
        obj_name = Mid(FileName, 1, InStrRev(FileName, ".") - 1)
        Kill Path & obj_name & "." & Ext
        Name Path & obj_name & ".sanitize" As Path & obj_name & "." & Ext
        FileName = Dir()
    Loop
End Sub
' Import References from a CSV, true=SUCCESS
Private Function ImportReferences(obj_path As String) As Boolean
    Dim fso, InFile
    Dim line As String
    Dim item() As String
    Dim GUID As String
    Dim Major As Long
    Dim Minor As Long
    Dim FileName As String
    FileName = Dir(obj_path & "references.csv")
    If Len(FileName) = 0 Then
        ImportReferences = False
        Exit Function
    End If
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set InFile = fso.OpenTextFile(FileName, ForReading)
    Do Until InFile.AtEndOfStream
        line = InFile.ReadLine
        item = Split(line, ",")
        GUID = item(0)
        Major = Clong(item(1))
        Minor = Clong(item(2))
        Application.References.AddFromGuid GUID, Major, Minor
    Loop
    Close InFile
    Set InFile = Nothing
    Set fso = Nothing
    ImportReferences = True
End Function
' Export References to a CSV
Private Sub ExportReferences(obj_path As String)
    Dim fso, OutFile
    Dim line As String
    Dim ref As Reference
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set OutFile = fso.CreateTextFile(obj_path & "references.csv", True)
    For Each ref In Application.References
        line = ref.GUID & "," & CStr(ref.Major) & "," & CStr(ref.Minor)
        OutFile.WriteLine line
    Next
    OutFile.Close
End Sub
' Save a Table Definition as SQL statement
Public Sub ExportTableDef(Db As Database, td As TableDef, tableName As String, FileName As String)
    Dim sql() As String
    Dim csql As String
    Dim idx As Index
    Dim fi As Field
    Dim i As Integer
    Dim nrSql As Integer
    Dim f As Field
    Dim rel As Relation
    Dim fso, OutFile
    Dim ff As Object
    'Debug.Print tableName
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set OutFile = fso.CreateTextFile(FileName, True)
    nrSql = 2
    ReDim Preserve sql(nrSql)
    sql(1) = "CREATE TABLE " & strName(tableName) & " ("

    For Each fi In td.Fields
        sql(0) = ""
        sql(1) = sql(1) & strName(fi.name) & " "
        If (fi.Attributes And dbAutoIncrField) Then
            sql(1) = sql(1) & "AUTOINCREMENT"
        Else
            sql(1) = sql(1) & strType(fi.Type) & " "
        End If
        Select Case fi.Type
            Case dbText, dbVarBinary
                sql(1) = sql(1) & "(" & fi.Size & ")"
            Case Else
        End Select
        For Each idx In td.Indexes
            If idx.Fields.count = 1 And idx.Fields(0).name = fi.name Then
                 
                If idx.Primary Then sql(0) = sql(0) & " PRIMARY KEY "
                If idx.Unique Then sql(0) = sql(0) & " UNIQUE "
                If idx.Required Then sql(0) = sql(0) & " NOT NULL "
                '
                If idx.Foreign Then
                Set ff = idx.Fields
                sql(0) = sql(0) & formatReferences(Db, ff, tableName)
                '
                End If
                If Len(sql(0)) > 0 Then sql(0) = " CONSTRAINT " & idx.name & sql(0)
            End If
        Next
        sql(1) = sql(1) + sql(0)
        sql(1) = sql(1) + ", "
    Next
    
    For Each idx In td.Indexes
        If idx.Fields.count > 1 Then
            If Len(sql(1)) = 0 Then sql(1) = sql(1) & " CONSTRAINT " & idx.name
            sql(1) = sql(1) & formatConstraint(idx.Primary, "PRIMARY KEY", idx)
            sql(1) = sql(1) & formatConstraint(idx.Unique, "UNIQUE", idx)
            sql(1) = sql(1) & formatConstraint(idx.Required, "NOT NULL", idx)
            sql(0) = ""
            sql(0) = formatConstraint(idx.Foreign, "FOREIGN KEY", idx)
            If Len(sql(0)) > 0 Then
                sql(1) = sql(1) & sql(0)
                sql(1) = sql(1) & formatReferences(Db, idx.Fields, tableName)
            End If
        End If
    Next
    sql(1) = Left(sql(1), Len(sql(1)) - 2) & ")"

    'Debug.Print sql
    OutFile.WriteLine sql(1)
    
    OutFile.Close
End Sub
Private Function formatReferences(Db As Database, ff As Object, tableName As String)
    Dim rel As Relation
    Dim sql As String
    Dim f As Field
    For Each rel In Db.Relations
        If (rel.foreignTable = tableName) Then
         If FieldsIdentical(ff, rel.Fields) Then
          sql = " REFERENCES "
          sql = sql & rel.Table & " ("
          For Each f In rel.Fields
            sql = sql & strName(f.name) & ","
          Next
          sql = Left(sql, Len(sql) - 1) & ")"
          If rel.Attributes And dbRelationUpdateCascade Then
            sql = sql + " ON UPDATE CASCADE "
          End If
          If rel.Attributes And dbRelationDeleteCascade Then
            sql = sql + " ON DELETE CASCADE "
          End If
          Exit For
         End If
        End If
    Next
    formatReferences = sql
End Function
Private Function formatConstraint(isConstraint As Boolean, keyw As String, idx As Index) As String
    Dim sql As String
    Dim fi As Field
    
    sql = sql & strName(idx.name)
    If isConstraint Then
        sql = sql & " " & keyw & " ("
        For Each fi In idx.Fields
            sql = sql & strName(fi.name) & ","
        Next
        sql = Left(sql, Len(sql) - 1) & ")"
        formatConstraint = sql
    Else
        formatConstraint = ""
    End If
            
End Function

Private Function strName(s As String) As String
    If InStr(s, " ") > 0 Then
        strName = "[" & s & "]"
    ElseIf UCase(s) = "UNIQUE" Then
        strName = "[" & s & "]"
    Else
        strName = s
    End If
End Function
Private Function strType(i As Integer) As String
    Select Case i
    Case dbLongBinary
        strType = "LONGBINARY"
    Case dbBinary
        strType = "BINARY"
    'Case dbBit missing enum
    '    strType = "BIT"
    Case dbAutoIncrField
        strType = "COUNTER"
    Case dbCurrency
        strType = "CURRENCY"
    Case dbDate, dbTime
        strType = "DATETIME"
    Case dbGUID
        strType = "GUID"
    Case dbMemo
        strType = "LONGTEXT"
    Case dbDouble
        strType = "DOUBLE"
    Case dbSingle
        strType = "SINGLE"
    Case dbByte
        strType = "UNSIGNED BYTE"
    Case dbInteger
        strType = "SHORT"
    Case dbLong
        strType = "LONG"
    Case dbNumeric
        strType = "NUMERIC"
    Case dbText
        strType = "VARCHAR"
    Case Else
        strType = "VARCHAR"
    End Select
End Function
Private Function FieldsIdentical(ff As Object, gg As Object) As Boolean
    Dim f As Field
    If ff.count <> gg.count Then
        FieldsIdentical = False
        Exit Function
    End If
    For Each f In ff
        If Not FieldInFields(f, gg) Then
        FieldsIdentical = False
        Exit Function
        End If
    Next
    FieldsIdentical = True
        
    
End Function

Private Function FieldInFields(fi As Field, ff As Fields) As Boolean
    Dim f As Field
    For Each f In ff
        If f.name = fi.name Then
            FieldInFields = True
            Exit Function
        End If
    Next
    FieldInFields = False
End Function

' Determine if a table or exists.
' based on sample code of support.microsoftcom
' ARGUMENTS:
'    TName: The name of a table or query.
'
' RETURNS: True (it exists) or False (it does not exist).
Function TableExists(TName As String) As Boolean
        Dim Db As Database, Found As Boolean, Test As String
        Const NAME_NOT_IN_COLLECTION = 3265

         ' Assume the table or query does not exist.
        Found = False
        Set Db = CurrentDb()

         ' Trap for any errors.
        On Error Resume Next
         
         ' See if the name is in the Tables collection.
        Test = Db.TableDefs(TName).name
        If Err <> NAME_NOT_IN_COLLECTION Then Found = True

        ' Reset the error variable.
        Err = 0

        TableExists = Found

End Function

' Main entry point for EXPORT. Export all forms, reports, queries,
' macros, modules, and lookup tables to `source` folder under the
' database's folder.
Public Sub ExportAllSource()
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
    Dim ucs2 As Boolean
    Dim tblName As Variant

    Set Db = CurrentDb

    CloseFormsReports
    InitUsingUcs2

    source_path = ProjectPath() & "source\"
    MkDirIfNotExist source_path

    Debug.Print

    obj_path = source_path & "queries\"
    ClearTextFilesFromDir obj_path, "bas"
    Debug.Print PadRight("Exporting queries...", 24);
    obj_count = 0
    For Each qry In Db.QueryDefs
        If Left(qry.name, 1) <> "~" Then
            ExportObject acQuery, qry.name, obj_path & qry.name & ".bas", UsingUcs2
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
            If Len(Dir(obj_path & tblName & ".txt")) > 0 Then
                obj_count = obj_count + 1
            End If
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
        For Each doc In Db.Containers(obj_type_name).Documents
            If Left(doc.name, 1) <> "~" Then
                If obj_type_label = "modules" Then
                    ucs2 = False
                Else
                    ucs2 = UsingUcs2
                End If
                ExportObject obj_type_num, doc.name, obj_path & doc.name & ".bas", ucs2
                obj_count = obj_count + 1
            End If
        Next
        Debug.Print "[" & obj_count & "]"

        If obj_type_label <> "modules" Then
            SanitizeTextFiles obj_path, "bas"
        Else
            ' Make sure all modules find their needed references
            ExportReferences obj_path
        End If
    Next

    DelIfExist TempFile()
    
    Dim td As TableDef
    Dim tds As TableDefs
    Set tds = Db.TableDefs

    obj_type_label = "tbldef"
    obj_type_name = "Table_Def"
    obj_type_num = acTable
    obj_path = source_path & obj_type_label & "\"
    obj_count = 0
    MkDirIfNotExist Left(obj_path, InStrRev(obj_path, "\"))
    ClearTextFilesFromDir obj_path, "def"
    Debug.Print PadRight("Exporting " & obj_type_label & "...", 24);
    
    For Each td In tds
        ' This is not a system table
        ' this is not a temporary table
        ' this is not an external table
        If Left$(td.name, 4) <> "MSys" And _
        Left(td.name, 1) <> "~" _
        And Len(td.Connect) = 0 _
        Then
            'Debug.Print
            ExportTableDef Db, td, td.name, obj_path & td.name & ".sql"
            obj_count = obj_count + 1
        End If
    Next
    Debug.Print "[" & obj_count & "]"
    
    
    Debug.Print "Done."
End Sub

' Main entry point for IMPORT. Import all forms, reports, queries,
' macros, modules, and lookup tables from `source` folder under the
' database's folder.
Public Sub ImportAllSource()
    Dim Db As Object ' DAO.Database
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
    Dim FileName As String
    Dim obj_name As String
    Dim ucs2 As Boolean

    Set Db = CurrentDb
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
    FileName = Dir(obj_path & "*.bas")
    If Len(FileName) > 0 Then
        Debug.Print PadRight("Importing queries...", 24);
        obj_count = 0
        Do Until Len(FileName) = 0
            obj_name = Mid(FileName, 1, InStrRev(FileName, ".") - 1)
            ImportObject acQuery, obj_name, obj_path & FileName, UsingUcs2
            obj_count = obj_count + 1
            FileName = Dir()
        Loop
        Debug.Print "[" & obj_count & "]"
    End If

    ' restore table definitions
    obj_path = source_path & "tbldef\"
    FileName = Dir(obj_path & "*.sql")
    If Len(FileName) > 0 Then
        Debug.Print PadRight("Importing tabledefs...", 24);
        obj_count = 0
        Do Until Len(FileName) = 0
            obj_name = Mid(FileName, 1, InStrRev(FileName, ".") - 1)
            ImportTableDef CStr(obj_name), obj_path & FileName
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
            obj_name = Mid(FileName, 1, InStrRev(FileName, ".") - 1)
            ImportTable CStr(obj_name), obj_path
            obj_count = obj_count + 1
            FileName = Dir()
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
        FileName = Dir(obj_path & "*.bas")
        If Len(FileName) > 0 Then
            Debug.Print PadRight("Importing " & obj_type_label & "...", 24);
            obj_count = 0
            Do Until Len(FileName) = 0
                obj_name = Mid(FileName, 1, InStrRev(FileName, ".") - 1)
                If obj_name <> "AppCodeImportExport" Then
                    If obj_type_label = "modules" Then
                        ucs2 = False
                    Else
                        ucs2 = UsingUcs2
                    End If
                    ImportObject obj_type_num, obj_name, obj_path & FileName, ucs2
                    obj_count = obj_count + 1
                End If
                FileName = Dir()
            Loop
            Debug.Print "[" & obj_count & "]"
            If obj_type_label = "modules" Then
                If Not ImportReferences(obj_path) Then
                    Debug.Print
                    Debug.Print "Info: no references file in " & obj_path
                End If
            End If
        End If
    Next

    DelIfExist TempFile()
    Dim td As TableDef
    Debug.Print "Done."
End Sub

' Build SQL to export `tbl_name` sorted by each field from first to last
Private Function TableExportSql(tbl_name As String)
    Dim rs As Object ' DAO.Recordset
    Dim fieldObj As Object ' DAO.Field
    Dim sb() As String, count As Integer

    Set rs = CurrentDb.OpenRecordset(tbl_name)
    
    sb = Sb_Init()
    Sb_Append sb, "SELECT "
    count = 0
    For Each fieldObj In rs.Fields
        If count > 0 Then Sb_Append sb, ", "
        Sb_Append sb, "[" & fieldObj.name & "]"
        count = count + 1
    Next
    Sb_Append sb, " FROM [" & tbl_name & "] ORDER BY "
    count = 0
    For Each fieldObj In rs.Fields
        If count > 0 Then Sb_Append sb, ", "
        Sb_Append sb, "[" & fieldObj.name & "]"
        count = count + 1
    Next

    TableExportSql = Sb_Get(sb)

End Function

' Export the lookup table `tblName` to `source\tables`.
Private Sub ExportTable(tbl_name As String, obj_path As String)
    Dim fso, OutFile
    Dim rs As Recordset ' DAO.Recordset
    Dim fieldObj As Object ' DAO.Field
    Dim C As Long, Value As Variant
    ' Checks first
    If Not TableExists(tbl_name) Then
        Debug.Print "Error: Table " & tbl_name & " missing"
        Exit Sub
    End If
    Set rs = CurrentDb.OpenRecordset(TableExportSql(tbl_name))
    If rs.RecordCount = 0 Then
        Debug.Print "Error: Table " & tbl_name & "  empty"
        rs.Close
        Exit Sub
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    ' open file for writing with Create=True, Unicode=True (USC-2 Little Endian format)
    MkDirIfNotExist obj_path
    Set OutFile = fso.CreateTextFile(TempFile(), True, True)

    C = 0
    For Each fieldObj In rs.Fields
        If C <> 0 Then OutFile.write vbTab
        C = C + 1
        OutFile.write fieldObj.name
    Next
    OutFile.write vbCrLf

    rs.MoveFirst
    Do Until rs.EOF
        C = 0
        For Each fieldObj In rs.Fields
            If C <> 0 Then OutFile.write vbTab
            C = C + 1
            Value = rs(fieldObj.name)
            If IsNull(Value) Then
                Value = ""
            Else
                Value = Replace(Value, "\", "\\")
                Value = Replace(Value, vbCrLf, "\n")
                Value = Replace(Value, vbCr, "\n")
                Value = Replace(Value, vbLf, "\n")
                Value = Replace(Value, vbTab, "\t")
            End If
            OutFile.write Value
        Next
        OutFile.write vbCrLf
        rs.MoveNext
    Loop
    rs.Close
    OutFile.Close

    ConvertUcs2Utf8 TempFile(), obj_path & tbl_name & ".txt"
End Sub
' Kill Table if Exists
Private Sub KillTable(tableName As String)
    If TableExists(tblName) Then
        Db.Execute "DROP TABLE [" & tblName & "]"
    End If
End Sub
' Import Table Definition
Private Sub ImportTableDef(tblName As String, FilePath As String)
    Dim Db As Object ' DAO.Database
    Dim fso, InFile As Object
    Dim buf As String
    Dim p As Integer
    Dim p1 As Integer
    Dim strFields() As String
    Dim strRef As String
    Dim strMsg As String
    Dim strForeignKeys() As String
    Dim s
    Dim n As Integer
    Dim i As Integer
    Dim j As Integer
    n = -1
    Set fso = CreateObject("Scripting.FileSystemObject")
    ConvertUtf8Ucs2 FilePath, TempFile()
    ' open file for reading with Create=False, Unicode=True (USC-2 Little Endian format)
    Set InFile = fso.OpenTextFile(TempFile(), ForReading, False, TristateTrue)
    Set Db = CurrentDb
    KillTable tblName
    buf = InFile.ReadLine()
    Do Until InFile.AtEndOfStream
        buf = buf & InFile.ReadLine()
    Loop
    
    ' The following block is needed because "on update" actions may cause problems
    For Each s In Split("UPDATE|DELETE", "|")
    p = InStr(buf, "ON " & s & " CASCADE")
    While p > 0
        n = n + 1
        ReDim Preserve K(n)
        K(n).Table = tblName
        K(n).isUpdate = (s = "UPDATE")
        
        buf = Left(buf, p - 1) & Mid(buf, p + 18)
        p = InStrRev(buf, "REFERENCES", p)
        p1 = InStr(p, buf, "(")
        K(n).foreignFields = Split(SubString(p1, buf, "(", ")"), ",")
        K(n).foreignTable = Trim(Mid(buf, p + 10, p1 - p - 10))
        p = InStrRev(buf, "CONSTRAINT", p1)
        p1 = InStrRev(buf, "FOREIGN KEY", p1)
        If (p1 > 0) And (p > 0) And (p1 > p) Then
        ' multifield index
            K(n).refFields = Split(SubString(p1, buf, "(", ")"), ",")
        ElseIf p1 = 0 Then
        ' single field
        End If
        p = InStr(p, "ON " & s & " CASCADE", buf)
    Wend
    Next
    On Error Resume Next
    For i = 0 To n
        strMsg = K(i).Table & " to " & K(i).foreignTable
        strMsg = strMsg & "(  "
        For j = 0 To UBound(K(i).refFields)
            strMsg = strMsg & K(i).refFields(j) & ", "
        Next j
        strMsg = Left(strMsg, Len(strMsg) - 2) & ") to ("
        For j = 0 To UBound(K(i).foreignFields)
            strMsg = strMsg & K(i).foreignFields(j) & ", "
        Next j
        strMsg = Left(strMsg, Len(strMsg) - 2) & ") Check "
        If K(i).isUpdate Then
            strMsg = strMsg & " on update cascade " & vbCrLf
        Else
            strMsg = strMsg & " on delete cascade " & vbCrLf
        End If
    Next
    On Error GoTo 0
    Db.Execute buf
    InFile.Close
    If Len(strMsg) > 0 Then MsgBox strMsg, vbOKOnly, "Correct manually"
End Sub
' returns substring between e.g. "(" and ")", internal brackets ar skippped
Public Function SubString(p As Integer, s As String, startsWith As String, endsWith As String)
    Dim start As Integer
    Dim last As Integer
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
' Import the lookup table `tblName` from `source\tables`.
Private Sub ImportTable(tblName As String, obj_path As String)
    Dim Db As Object ' DAO.Database
    Dim rs As Object ' DAO.Recordset
    Dim fieldObj As Object ' DAO.Field
    Dim fso, InFile As Object
    Dim C As Long, buf As String, Values() As String, Value As Variant

    Set fso = CreateObject("Scripting.FileSystemObject")
    ConvertUtf8Ucs2 obj_path & tblName & ".txt", TempFile()
    ' open file for reading with Create=False, Unicode=True (USC-2 Little Endian format)
    Set InFile = fso.OpenTextFile(TempFile(), ForReading, False, TristateTrue)
    Set Db = CurrentDb

    Db.Execute "DELETE FROM [" & tblName & "]"
    Set rs = Db.OpenRecordset(tblName)
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
                rs(fieldObj.name) = Value
                C = C + 1
            Next
            rs.Update
        End If
    Loop

    rs.Close
    InFile.Close
End Sub
' Expose for use as function, can be called by query
Public Function make()
    ImportAllSource
End Function