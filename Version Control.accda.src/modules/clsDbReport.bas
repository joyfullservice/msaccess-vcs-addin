Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : This class extends the IDbComponent class to perform the specific
'           : operations required by this particular object type.
'           : (I.e. The specific way you export or import this component.)
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit

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


Private m_Report As AccessObject
Private m_Options As clsOptions
Private m_AllItems As Collection

' This requires us to use all the public methods and properties of the implemented class
' which keeps all the component classes consistent in how they are used in the export
' and import process. The implemented functions should be kept private as they are called
' from the implementing class, not this class.
Implements IDbComponent


'---------------------------------------------------------------------------------------
' Procedure : Export
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Export the individual database component (table, form, query, etc...)
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Export()
    
    Dim strFile As String
    
    ' Export main report object
    SaveComponentAsText acReport, m_Report.Name, IDbComponent_SourceFile, IDbComponent_Options
    
    ' Export print vars if selected
    If IDbComponent_Options.SavePrintVars Then
        strFile = IDbComponent_BaseFolder & GetSafeFileName(m_Report.Name) & ".json"
        ExportPrintVars m_Report.Name, strFile
    End If
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ExportPrintVars
' Author    : Adam Waller
' Date      : 1/25/2019
' Purpose   : Exports print vars for reports
'           : https://docs.microsoft.com/en-us/office/vba/api/access.report.prtdevmode
'---------------------------------------------------------------------------------------
'
Public Sub ExportPrintVars(strReport As String, strFile As String)

    Dim DevModeString As str_DEVMODE
    Dim DevModeExtra As String
    Dim DM As type_DEVMODE
    Dim rpt As Report
    Dim dItems As Scripting.Dictionary

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

        Set dItems = New Scripting.Dictionary
        With dItems
            .Add "Orientation", DM.intOrientation
            .Add "PaperSize", DM.intPaperSize
            .Add "PaperLength", DM.intPaperLength
            .Add "PaperWidth", DM.intPaperWidth
            .Add "Scale", DM.intScale
        End With

        ' Write output to file
        WriteJsonFile Me, dItems, strFile, "Report Print Settings"

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


'---------------------------------------------------------------------------------------
' Procedure : Import
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Import the individual database component from a file.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Import(strFile As String)

End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetAllFromDB
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a collection of class objects represented by this component type.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_GetAllFromDB(Optional cOptions As clsOptions) As Collection
    
    Dim rpt As AccessObject
    Dim cReport As IDbComponent

    ' Build collection if not already cached
    If m_AllItems Is Nothing Then
    
        ' Use parameter options if provided.
        If Not cOptions Is Nothing Then Set IDbComponent_Options = cOptions
    
        Set m_AllItems = New Collection
        For Each rpt In CurrentProject.AllReports
            Set cReport = New clsDbReport
            Set cReport.DbObject = rpt
            Set cReport.Options = IDbComponent_Options
            m_AllItems.Add cReport, rpt.Name
        Next rpt
    End If

    ' Return cached collection
    Set IDbComponent_GetAllFromDB = m_AllItems
        
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetFileList
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a list of file names to import for this component type.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_GetFileList() As Collection
    Set IDbComponent_GetFileList = GetFilePathsInFolder(IDbComponent_BaseFolder & "*.bas")
End Function


'---------------------------------------------------------------------------------------
' Procedure : ClearOrphanedSourceFiles
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Remove any source files for objects not in the current database.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_ClearOrphanedSourceFiles() As Variant
    ClearFilesByExtension IDbComponent_BaseFolder, "pv"
    If Not IDbComponent_Options.SavePrintVars Then ClearFilesByExtension IDbComponent_BaseFolder, "json"
    ClearOrphanedSourceFiles Me, "bas", "json"
End Function


'---------------------------------------------------------------------------------------
' Procedure : DateModified
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : The date/time the object was modified. (If possible to retrieve)
'           : If the modified date cannot be determined (such as application
'           : properties) then this function will return 0.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_DateModified() As Date
    IDbComponent_DateModified = m_Report.DateModified
End Function


'---------------------------------------------------------------------------------------
' Procedure : SourceModified
' Author    : Adam Waller
' Date      : 4/27/2020
' Purpose   : The date/time the source object was modified. In most cases, this would
'           : be the date/time of the source file, but it some cases like SQL objects
'           : the date can be determined through other means, so this function
'           : allows either approach to be taken.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_SourceModified() As Date
    If FSO.FileExists(IDbComponent_SourceFile) Then IDbComponent_SourceModified = FileDateTime(IDbComponent_SourceFile)
End Function


'---------------------------------------------------------------------------------------
' Procedure : Category
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a category name for this type. (I.e. forms, queries, macros)
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_Category() As String
    IDbComponent_Category = "reports"
End Property


'---------------------------------------------------------------------------------------
' Procedure : BaseFolder
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return the base folder for import/export of this component.
'---------------------------------------------------------------------------------------
Private Property Get IDbComponent_BaseFolder() As String
    IDbComponent_BaseFolder = IDbComponent_Options.GetExportFolder & "reports\"
End Property


'---------------------------------------------------------------------------------------
' Procedure : Name
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a name to reference the object for use in logs and screen output.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_Name() As String
    IDbComponent_Name = m_Report.Name
End Property


'---------------------------------------------------------------------------------------
' Procedure : SourceFile
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return the full path of the source file for the current object.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_SourceFile() As String
    IDbComponent_SourceFile = IDbComponent_BaseFolder & GetSafeFileName(m_Report.Name) & ".bas"
End Property


'---------------------------------------------------------------------------------------
' Procedure : Count
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a count of how many items are in this category.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_Count() As Long
    IDbComponent_Count = IDbComponent_GetAllFromDB.Count
End Property


'---------------------------------------------------------------------------------------
' Procedure : ComponentType
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : The type of component represented by this class.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_ComponentType() As eDatabaseComponentType
    IDbComponent_ComponentType = edbReport
End Property


'---------------------------------------------------------------------------------------
' Procedure : Upgrade
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Run any version specific upgrade processes before importing.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Upgrade()
    ' No upgrade needed.
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Options
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return or set the options being used in this context.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_Options() As clsOptions
    If m_Options Is Nothing Then Set m_Options = LoadOptions
    Set IDbComponent_Options = m_Options
End Property
Private Property Set IDbComponent_Options(ByVal RHS As clsOptions)
    Set m_Options = RHS
End Property


'---------------------------------------------------------------------------------------
' Procedure : DbObject
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : This represents the database object we are dealing with.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_DbObject() As Object
    Set IDbComponent_DbObject = m_Report
End Property
Private Property Set IDbComponent_DbObject(ByVal RHS As Object)
    Set m_Report = RHS
End Property


'---------------------------------------------------------------------------------------
' Procedure : SingleFile
' Author    : Adam Waller
' Date      : 4/24/2020
' Purpose   : Returns true if the export of all items is done as a single file instead
'           : of individual files for each component. (I.e. properties, references)
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_SingleFile() As Boolean
End Property


'---------------------------------------------------------------------------------------
' Procedure : Parent
' Author    : Adam Waller
' Date      : 4/24/2020
' Purpose   : Return a reference to this class as an IDbComponent. This allows you
'           : to reference the public methods of the parent class without needing
'           : to create a new class object.
'---------------------------------------------------------------------------------------
'
Public Property Get Parent() As IDbComponent
    Set Parent = Me
End Property