Attribute VB_Name = "VCS_Reference"
Option Compare Database

Option Private Module
Option Explicit

Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const TristateTrue = -1, TristateFalse = 0, TristateUseDefault = -2

' Import References from a CSV, true=SUCCESS
Public Function ImportReferences(obj_path As String) As Boolean
    Dim FSO, InFile
    Dim line As String
    Dim item() As String
    Dim GUID As String
    Dim Major As Long
    Dim Minor As Long
    Dim fileName As String
    Dim refName As String
    fileName = Dir(obj_path & "references.csv")
    If Len(fileName) = 0 Then
        ImportReferences = False
        Exit Function
    End If
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set InFile = FSO.OpenTextFile(obj_path & fileName, ForReading)
On Error GoTo failed_guid
    Do Until InFile.AtEndOfStream
        line = InFile.ReadLine
        item = Split(line, ",")
        If UBound(item) = 2 Then 'a ref with a guid
          GUID = Trim(item(0))
          Major = CLng(item(1))
          Minor = CLng(item(2))
          Application.References.AddFromGuid GUID, Major, Minor
        Else
          refName = Trim(item(0))
          Application.References.AddFromFile refName
        End If
go_on:
    Loop
On Error GoTo 0
    InFile.Close
    Set InFile = Nothing
    Set FSO = Nothing
    ImportReferences = True
    Exit Function
failed_guid:
    If Err.Number = 32813 Then
        'The reference is already present in the access project - so we can ignore the error
        Resume Next
    Else
        MsgBox "Failed to register " & GUID, , "Error: " & Err.Number
        'Do we really want to carry on the import with missing references??? - Surely this is fatal
        Resume go_on
    End If
    
End Function

' Export References to a CSV
Public Sub ExportReferences(obj_path As String)
    Dim FSO, OutFile
    Dim line As String
    Dim ref As Reference

    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set OutFile = FSO.CreateTextFile(obj_path & "references.csv", True)
    For Each ref In Application.References
        If ref.GUID > "" Then ' references of types mdb,accdb,mde etc don't have a GUID
            If Not ref.BuiltIn Then
                line = ref.GUID & "," & CStr(ref.Major) & "," & CStr(ref.Minor)
                OutFile.WriteLine line
            End If
        Else
            line = ref.FullPath
            OutFile.WriteLine line
        End If
    Next
    OutFile.Close
End Sub