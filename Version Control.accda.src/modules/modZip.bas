Attribute VB_Name = "modZip"
'---------------------------------------------------------------------------------------
' Module    : modZip
' Author    : Adam Waller
' Date      : 12/4/2020
' Purpose   : Functions for creating and working with Zip files
'---------------------------------------------------------------------------------------
Option Compare Database
Option Private Module
Option Explicit


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
    If blnPauseTillFinished Then
        sngTimeout = Timer + intTimeoutSeconds
        Do While Timer < sngTimeout
            ' Check to see if all the items have been copied.
            If oApp.Namespace(varZip).Items.Count = lngCount Then Exit Do
            Pause 0.5
        Loop
    End If
    
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
    varFolder = strFolder & PathSep

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

