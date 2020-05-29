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

' See the following links for additional technical details regarding the DEVMODE strcture:
' https://docs.microsoft.com/en-us/office/vba/api/access.report.prtdevmode
' https://stackoverflow.com/questions/49560317/64-bit-word-vba-devmode-dmduplex-returns-4
' http://toddmcdermid.blogspot.com/2009/02/microsoft-access-2003-and-printer.html

Private Type str_DEVMODE
    RGB As String * 220
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
    intUnusedPadding As Integer
    intBitsPerPel As Integer
    lngPelsWidth As Long
    lngPelsHeight As Long
    lngDisplayFlags As Long
    lngDisplayFrequency As Long
    lngICMMethod As Long
    lngICMIntent As Long
    lngMediaType As Long
    lngDitherType As Long
    lngReserved1 As Long
    lngReserved2 As Long
End Type

Private m_Report As AccessObject
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

    ' Export main report object
    SaveComponentAsText acReport, m_Report.Name, IDbComponent_SourceFile
    
    ' Export print vars if selected
    If Options.SavePrintVars Then
        ExportPrintVars m_Report.Name, GetPrintVarsFileName(m_Report.Name)
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

        ' Read report devmode into structure and convert to dictionary
        DevModeExtra = rpt.PrtDevMode
        DevModeString.RGB = DevModeExtra
        LSet DM = DevModeString
        Set dItems = DevModeToDictionary(DM)

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


'---------------------------------------------------------------------------------------
' Procedure : ImportPrintVars
' Author    : Adam Waller
' Date      : 5/7/2020
' Purpose   : Import the print vars back into the report.
'---------------------------------------------------------------------------------------
'
Public Sub ImportPrintVars(strFile As String)
    
'    Dim DevModeString As str_DEVMODE
'    Dim tDevMode As type_DEVMODE
'    Dim DevModeExtra As String
'    Dim dFile As Dictionary
'    Dim strReport As String
'
'    Set dFile = ReadJsonFile(strFile)
'    If Not dFile Is Nothing Then
'
'        ' Prepare data structures
'        tDevMode = DictionaryToDevMode(dFile("Items"))
'        LSet DevModeString = tDevMode
'        Mid(DevModeExtra, 1, 94) = DevModeString.RGB
'
'        ' Apply to report
'        strReport = GetObjectNameFromFileName(strFile)
'        DoCmd.Echo False
'        DoCmd.OpenReport strReport, acViewDesign
'        Reports(strReport).PrtDevMode = DevModeExtra
'        DoCmd.Close acReport, strReport, acSaveYes
'    End If
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Import
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Import the individual database component from a file.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Import(strFile As String)
    
    Dim strReport As String
    
    ' Import the report object
    strReport = GetObjectNameFromFileName(strFile)
    LoadComponentFromText acReport, strReport, strFile
    
    ' Import the print vars if specified
    If Options.SavePrintVars Then
        ImportPrintVars GetPrintVarsFileName(strReport)
    End If
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetAllFromDB
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a collection of class objects represented by this component type.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_GetAllFromDB() As Collection
    
    Dim rpt As AccessObject
    Dim cReport As IDbComponent

    ' Build collection if not already cached
    If m_AllItems Is Nothing Then
        Set m_AllItems = New Collection
        For Each rpt In CurrentProject.AllReports
            Set cReport = New clsDbReport
            Set cReport.DbObject = rpt
            m_AllItems.Add cReport, rpt.Name
        Next rpt
    End If

    ' Return cached collection
    Set IDbComponent_GetAllFromDB = m_AllItems
        
End Function


'---------------------------------------------------------------------------------------
' Procedure : DevModeToDictionary
' Author    : Adam Waller
' Date      : 5/7/2020
' Purpose   : Convert a DEVMODE type to a dictionary.
'---------------------------------------------------------------------------------------
'
Private Function DevModeToDictionary(cDev As type_DEVMODE) As Dictionary
    Set DevModeToDictionary = New Dictionary
    With DevModeToDictionary
        .Add "DeviceName", cDev.strDeviceName
        .Add "SpecVersion", cDev.intSpecVersion
        .Add "DriverVersion", cDev.intDriverVersion
        .Add "Size", cDev.intSize
        .Add "DriverExtra", cDev.intDriverExtra
        .Add "Fields", cDev.lngFields
        .Add "Orientation", cDev.intOrientation
        .Add "PaperSize", cDev.intPaperSize
        .Add "PaperLength", cDev.intPaperLength
        .Add "PaperWidth", cDev.intPaperWidth
        .Add "Scale", cDev.intScale
        .Add "Copies", cDev.intCopies
        .Add "DefaultSource", cDev.intDefaultSource
        .Add "PrintQuality", cDev.intPrintQuality
        .Add "Color", cDev.intColor
        .Add "Duplex", cDev.intDuplex
        .Add "Resolution", cDev.intResolution
        .Add "TTOption", cDev.intTTOption
        .Add "Collate", cDev.intCollate
        .Add "FormName", cDev.strFormName
        .Add "UnusedPadding", cDev.intUnusedPadding
        .Add "BitsPerPel", cDev.intBitsPerPel
        .Add "PelsWidth", cDev.lngPelsWidth
        .Add "PelsHeight", cDev.lngPelsHeight
        .Add "DisplayFlags", cDev.lngDisplayFlags
        .Add "DisplayFrequency", cDev.lngDisplayFrequency
        .Add "ICMMethod", cDev.lngICMMethod
        .Add "ICMIntent", cDev.lngICMIntent
        .Add "MediaType", cDev.lngMediaType
        .Add "DitherType", cDev.lngDitherType
        .Add "Reserved1", cDev.lngReserved1
        .Add "Reserved2", cDev.lngReserved2
    End With
End Function


'---------------------------------------------------------------------------------------
' Procedure : DictionaryToDevMode
' Author    : Adam Waller
' Date      : 5/7/2020
' Purpose   : Excel formulas make it easy to edit these!
'---------------------------------------------------------------------------------------
'
Private Function DictionaryToDevMode(dDevMode As Dictionary) As type_DEVMODE
    With DictionaryToDevMode
        .strDeviceName = dDevMode("DeviceName")
        .intSpecVersion = dDevMode("SpecVersion")
        .intDriverVersion = dDevMode("DriverVersion")
        .intSize = dDevMode("Size")
        .intDriverExtra = dDevMode("DriverExtra")
        .lngFields = dDevMode("Fields")
        .intOrientation = dDevMode("Orientation")
        .intPaperSize = dDevMode("PaperSize")
        .intPaperLength = dDevMode("PaperLength")
        .intPaperWidth = dDevMode("PaperWidth")
        .intScale = dDevMode("Scale")
        .intCopies = dDevMode("Copies")
        .intDefaultSource = dDevMode("DefaultSource")
        .intPrintQuality = dDevMode("PrintQuality")
        .intColor = dDevMode("Color")
        .intDuplex = dDevMode("Duplex")
        .intResolution = dDevMode("Resolution")
        .intTTOption = dDevMode("TTOption")
        .intCollate = dDevMode("Collate")
        .strFormName = dDevMode("FormName")
        .intUnusedPadding = dDevMode("UnusedPadding")
        .intBitsPerPel = dDevMode("BitsPerPel")
        .lngPelsWidth = dDevMode("PelsWidth")
        .lngPelsHeight = dDevMode("PelsHeight")
        .lngDisplayFlags = dDevMode("DisplayFlags")
        .lngDisplayFrequency = dDevMode("DisplayFrequency")
        .lngICMMethod = dDevMode("ICMMethod")
        .lngICMIntent = dDevMode("ICMIntent")
        .lngMediaType = dDevMode("MediaType")
        .lngDitherType = dDevMode("DitherType")
        .lngReserved1 = dDevMode("Reserved1")
        .lngReserved2 = dDevMode("Reserved2")
    End With
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetPrintVarsFileName
' Author    : Adam Waller
' Date      : 5/7/2020
' Purpose   : Return the file name used to export/import print vars
'---------------------------------------------------------------------------------------
'
Private Function GetPrintVarsFileName(strReport As String) As String
    GetPrintVarsFileName = IDbComponent_BaseFolder & GetSafeFileName(strReport) & ".json"
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
Private Sub IDbComponent_ClearOrphanedSourceFiles()
    ClearFilesByExtension IDbComponent_BaseFolder, "pv"
    If Not Options.SavePrintVars Then ClearFilesByExtension IDbComponent_BaseFolder, "json"
    ClearOrphanedSourceFiles Me, "bas", "json"
End Sub


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
    IDbComponent_BaseFolder = Options.GetExportFolder & "reports\"
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
    IDbComponent_SingleFile = False
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