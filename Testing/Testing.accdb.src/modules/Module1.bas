Attribute VB_Name = "Module1"
Option Compare Database
Option Explicit


Private Type str_PRTMIP
    strRGB As String * 28
End Type


Private Type tDevNamesBuffer
    strBuffer As String * 255
End Type

Private Type type_PRTMIP
    xLeftMargin As Long
    yTopMargin As Long
    xRightMargin As Long
    yBotMargin As Long
    fDataOnly As Long
    xWidth As Long
    yHeight As Long
    fDefaultSize As Long
    cxColumns As Long
    yColumnSpacing As Long
    xRowSpacing As Long
    rItemLayout As Long
    fFastPrint As Long
    fDatasheet As Long
End Type


Private Type DEVNAMES
  wDriverOffset As Integer
  wDeviceOffset As Integer
  wOutputOffset As Integer
  wDefault As Integer
  extra(1 To 255) As Byte
End Type



Public Sub PrtMipCols(ByVal strName As String)

    Dim PrtMipString As str_PRTMIP
    Dim PM As type_PRTMIP
    Dim rpt As Report
    Const PM_HORIZONTALCOLS = 1953
    Const PM_VERTICALCOLS = 1954
    
    ' Open the report.
    DoCmd.OpenReport strName, acDesign
    Set rpt = Reports(strName)
    PrtMipString.strRGB = rpt.PrtMip
    LSet PM = PrtMipString
    
    ' Create two columns.
    PM.cxColumns = 2
    
    ' Set 0.25 inch between rows.
    PM.xRowSpacing = 0.25 * 1440
    
    ' Set 0.5 inch between columns.
    PM.yColumnSpacing = 0.5 * 1440
    PM.rItemLayout = PM_HORIZONTALCOLS
    
    ' Update property.
    LSet PrtMipString = PM
    rpt.PrtMip = PrtMipString.strRGB
    
    Set rpt = Nothing
    
End Sub


Public Sub TestPrinterMIP()

    Dim rpt As Report
    Dim tDevNames As DEVNAMES
    Dim tBuffer As tDevNamesBuffer
    Dim strData As String
    Dim lngNull As Long
    Dim lngStart As Long
    
    Set rpt = Report_rptDefaultPrinter
    Set rpt = Report_rptNavigationPaneGroups
    
    tBuffer.strBuffer = rpt.PrtDevNames
    LSet tDevNames = tBuffer
    
    ' Bytes in structure before the data string starts
    
    'debug.Print mid$(strdata,tDevNames.wDeviceOffset- lngstart,instr
    
    strData = StrConv(tDevNames.extra, vbUnicode)
    
    Debug.Print GetNullTermStringByOffset(strData, 7, tDevNames.wDriverOffset)
    Debug.Print GetNullTermStringByOffset(strData, 7, tDevNames.wDeviceOffset)
    Debug.Print GetNullTermStringByOffset(strData, 7, tDevNames.wOutputOffset)
    
    
    Stop
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetNullTermStringByOffset
' Author    : Adam Waller
' Date      : 5/18/2020
' Purpose   : Returns the value of a null-terminated string by offset.
'---------------------------------------------------------------------------------------
'
Private Function GetNullTermStringByOffset(strData As String, lngHeaderLen As Long, intOffset As Integer) As String
    
    Dim lngNull As Long
    Dim lngStart As Long
    
    lngStart = intOffset - lngHeaderLen
    lngNull = InStr(lngStart, strData, vbNullChar)
    
    ' Return the string if we found a null terminator
    If lngNull > 0 Then GetNullTermStringByOffset = Mid$(strData, lngStart, lngNull - lngStart)
    
End Function
