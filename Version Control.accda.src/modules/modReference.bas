Option Explicit
Option Compare Database
Option Private Module


' Import References from a CSV, true=SUCCESS
Public Function ImportReferences(obj_path As String) As Boolean
    Dim fso As New Scripting.FileSystemObject
    Dim InFile As Scripting.TextStream
    Dim line As String
    Dim Item() As String
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
    Set InFile = fso.OpenTextFile(obj_path & fileName, ForReading)
On Error GoTo failed_guid
    Do Until InFile.AtEndOfStream
        line = InFile.ReadLine
        Item = Split(line, ",")
        If UBound(Item) = 2 Then 'a ref with a guid
          GUID = Trim(Item(0))
          Major = CLng(Item(1))
          Minor = CLng(Item(2))
          Application.References.AddFromGuid GUID, Major, Minor
        Else
          refName = Trim(Item(0))
          Application.References.AddFromFile refName
        End If
go_on:
    Loop
On Error GoTo 0
    InFile.Close
    Set InFile = Nothing
    Set fso = Nothing
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
    
    Dim fso As New Scripting.FileSystemObject
    Dim OutFile As Scripting.TextStream
    Dim line As String
    Dim ref As Reference
    Dim obj_count As Integer
    
    Set OutFile = fso.CreateTextFile(obj_path & "references.csv", True)
    
    For Each ref In Application.References
        If ref.GUID > "" Then ' references of types mdb,accdb,mde etc don't have a GUID
          line = ref.GUID & "," & CStr(ref.Major) & "," & CStr(ref.Minor)
          OutFile.WriteLine line
        Else
          line = ref.FullPath
          OutFile.WriteLine line
        End If
        obj_count = obj_count + 1
        If ShowDebugInfo Then Debug.Print "  " & line
    Next ref
    OutFile.Close

    If ShowDebugInfo Then
        Debug.Print "[" & obj_count & "] references exported."
    Else
        Debug.Print "[" & obj_count & "]"
    End If
    
    Set OutFile = Nothing
    Set fso = Nothing
    
End Sub