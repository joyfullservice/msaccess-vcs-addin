Option Explicit
Option Compare Database
Option Private Module

' See https://docs.microsoft.com/en-us/office/vba/api/access.report.prtdevmode

Private Type str_DEVMODE
    RGB As String * 94
End Type

Private Type type_DEVMODE
    strDeviceName As String * 32
    intSpecVersion As Integer
    intDriverVersion As Integer
    intSize As Integer
    intDriverExtra As Integer
    lngFields As Long
    intOrientation As Integer
    intPaperSize As Integer
    intPaperLength As Integer
    intPaperWidth As Integer
    intScale As Integer
    intCopies As Integer
    intDefaultSource As Integer
    intPrintQuality As Integer
    intColor As Integer
    intDuplex As Integer
    intResolution As Integer
    intTTOption As Integer
    intCollate As Integer
    strFormName As String * 32
    lngPad As Long
    lngBits As Long
    lngPW As Long
    lngPH As Long
    lngDFI As Long
    lngDFr As Long
End Type


'---------------------------------------------------------------------------------------
' Procedure : ExportPrintVars
' Author    : Adam Waller
' Date      : 1/25/2019
' Purpose   : Exports print vars for reports
'           : https://docs.microsoft.com/en-us/office/vba/api/access.report.prtdevmode
'---------------------------------------------------------------------------------------
'
Public Sub ExportPrintVars(strReport As String, strFile As String, cOptions As clsOptions)
    
    Dim DevModeString As str_DEVMODE
    Dim DevModeExtra As String
    Dim DM As type_DEVMODE
    Dim rpt As Report
    Dim cData As New clsConcat
    
    'report must be open to access Report object
    'report must be opened in design view to save changes to the print vars
    Application.Echo False
    DoCmd.OpenReport strReport, acViewDesign
    Set rpt = Reports(strReport)
    rpt.Visible = False
    
    ' Make sure we don't have a null devmode
    If Not IsNull(rpt.PrtDevMode) Then
        
        ' Read report devmode into structure
        DevModeExtra = rpt.PrtDevMode
        DevModeString.RGB = DevModeExtra
        LSet DM = DevModeString
       
        ' Print out print var values
        With cData
            .Add "Orientation=":    .Add CStr(DM.intOrientation)
            .Add "PaperSize=":      .Add CStr(DM.intPaperSize)
            .Add "PaperLength=":    .Add CStr(DM.intPaperLength)
            .Add "PaperWidth=":     .Add CStr(DM.intPaperWidth)
            .Add "Scale=":          .Add CStr(DM.intScale)
        End With
       
        ' Write output to file
        WriteFile cData.GetStr, strFile
        
    Else
        ' DevMode was null
        Log.Add "  Warning: PrtDevMode is null"
    End If
        
    ' Clean up
    Set rpt = Nothing
    DoCmd.Close acReport, strReport, acSaveNo
    Application.Echo True
    
End Sub


Public Sub ImportPrintVars(obj_name As String, filePath As String)


 Dim DevModeString As str_DEVMODE
 Dim DevModeExtra As String
 Dim varLine As Variant
 
 Dim DM As type_DEVMODE
  Dim rpt As Report
 'report must be open to access Report object
 'report must be opened in design view to save changes to the print vars

  DoCmd.OpenReport obj_name, acViewDesign

 Set rpt = Reports(obj_name)

'read print vars into struct
 If Not IsNull(rpt.PrtDevMode) Then
    DevModeExtra = rpt.PrtDevMode
    DevModeString.RGB = DevModeExtra
    LSet DM = DevModeString
 Else
    Set rpt = Nothing
    DoCmd.Close acReport, obj_name, acSaveNo
    Debug.Print "Warning: PrtDevMode is null"
    Exit Sub
 End If
 
 Dim InFile As Scripting.TextStream ' Object
 Set InFile = FSO.OpenTextFile(filePath, ForReading)
 
 ' Loop through lines
 Do While Not InFile.AtEndOfStream
    varLine = Split(InFile.ReadLine, "=")
    If UBound(varLine) = 1 Then
        Select Case varLine(0)
            Case "Orientation":     DM.intOrientation = varLine(1)
            Case "PaperSize":       DM.intPaperSize = varLine(1)
            Case "PaperLength":     DM.intPaperLength = varLine(1)
            Case "PaperWidth":      DM.intPaperWidth = varLine(1)
            Case "Scale":           DM.intScale = varLine(1)
            Case Else
                Debug.Print "* Unknown print var: '" & varLine(0) & "'"
        End Select
    End If
 Loop
 
 InFile.Close
 
'write print vars back into report
LSet DevModeString = DM
 Mid(DevModeExtra, 1, 94) = DevModeString.RGB
 rpt.PrtDevMode = DevModeExtra

 Set rpt = Nothing

DoCmd.Close acReport, obj_name, acSaveYes

End Sub