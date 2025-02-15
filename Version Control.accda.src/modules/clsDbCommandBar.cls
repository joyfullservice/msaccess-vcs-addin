VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDbCommandBar"
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

Private m_TemplateCommandBar As CommandBar
Private m_CommandBar As CommandBar
Private m_Items(True To False) As Dictionary
Private m_dItems As Dictionary
Private m_FileList As Dictionary

Private Type udtThis
    BarProperties As Collection
    CtlReadProperties As Collection
    CtlWriteProperties As Collection
    dImages As Dictionary       ' Dictionary of image objects used with this CommandBar
End Type
Private this As udtThis

' This requires us to use all the public methods and properties of the implemented class
' which keeps all the component classes consistent in how they are used in the export
' and import process. The implemented functions should be kept private as they are called
' from the implementing class, not this class.
Implements IDbComponent


'---------------------------------------------------------------------------------------
' Procedure : Export
' Author    : Adam Waller
' Date      : 5/17/2021
' Purpose   : Export the individual database component (table, form, query, etc...)
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Export(Optional strAlternatePath As String)
    Dim strContent As String
    strContent = GetSource
    WriteFile strContent, Nz2(strAlternatePath, IDbComponent_SourceFile)
    SaveImages
    VCSIndex.Update Me, IIf(strAlternatePath = vbNullString, eatExport, eatAltExport), GetStringHash(strContent, True)
End Sub


'---------------------------------------------------------------------------------------
' Procedure : SaveImages
' Author    : Adam Waller
' Date      : 1/12/2024
' Purpose   : Saves the images for the menu bar
'---------------------------------------------------------------------------------------
'
Private Sub SaveImages()

    Dim varPic As Variant

    ' Make sure we actually have images to save
    If this.dImages Is Nothing Then Exit Sub
    If this.dImages.Count = 0 Then Exit Sub

    ' Verify the first path, to make sure the folder exists
    VerifyPath CStr(this.dImages.Keys(1))

    Perf.OperationStart "Save CommandBar Images"
    For Each varPic In this.dImages.Keys
        ' Save the item image to a file
        stdole.SavePicture this.dImages(varPic).Picture, varPic & "Picture.bmp"
        ' Also save image mask
        stdole.SavePicture this.dImages(varPic).Mask, varPic & "Mask.bmp"
    Next varPic
    Perf.OperationEnd

    ' After saving the images, we can clear the references
    ' to the image objects
    Set this.dImages = Nothing

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Import
' Author    : Adam Waller
' Date      : 5/17/2021
' Purpose   : Import the individual database component from a file.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Import(strFile As String)

    Dim dBar As Dictionary
    Dim strName As String

    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    ' Only import files with the correct extension.
    If Not strFile Like "*.json" Then Exit Sub
    Set dBar = ReadJsonFile(strFile)

    Perf.OperationStart "Building CommandBar"
    ' Delete any existing bar with this name.
    strName = GetObjectNameFromFileName(strFile)
    Set m_CommandBar = GetExisting(strName)
    If Not m_CommandBar Is Nothing Then
        If m_CommandBar.BuiltIn Then
            Log.Error eelError, "You cannot replace the existing built-in command bar: " & m_CommandBar.Name, ModuleName(Me) & ".Import"
            Exit Sub
        Else
            ' Delete the command bar
            m_CommandBar.Delete
            Set m_CommandBar = Nothing
        End If
    End If

    ' Now, create a new command bar, and add controls to it.
    Dim IsMenu As Boolean
    Dim Position As MsoBarPosition

    Position = -1 'set to invalid value to simulate the "missing" value which cannot be directly set in VBA...
    IsMenu = (dBar("Items").Item("Type") = msoBarTypeMenuBar)
    If CatchAny(eelNoError, vbNullString) Then
        IsMenu = False
    End If
    Position = dBar("Items").Item("Position")
    If CatchAny(eelNoError, vbNullString) Then
        Position = -1
    End If

    If Position = -1 Then
        Set m_CommandBar = CommandBars.Add(strName, , IsMenu, False)
    Else
        Set m_CommandBar = CommandBars.Add(strName, Position, IsMenu, False)
    End If
    BuildControls dBar("Items"), m_CommandBar
    Perf.OperationEnd

    ' Log any errors
    CatchAny eelError, "Importing CommandBar " & strName, ModuleName(Me) & ".Import"

    ' Save to index
    VCSIndex.Update Me, eatImport, GetDictionaryHash(GetDictionary)

End Sub


'---------------------------------------------------------------------------------------
' Procedure : BuildControls
' Author    : Adam Waller
' Date      : 1/13/2024
' Purpose   : Recursive function to build the controls on the popup menu
'---------------------------------------------------------------------------------------
'
Private Sub BuildControls(dParent As Dictionary, objItem As Object)

    Dim varProp As Variant
    Dim varCtl As Variant
    Dim dControl As Dictionary
    Dim varValue As Variant
    Dim varCtlProp As Variant
    Dim colItems As Collection
    Dim picItem As IPictureDisp
    Dim strPath As String

    LogUnhandledErrors
    On Error Resume Next

    If TypeOf objItem Is CommandBar Then
        ' Set command bar properties
        For Each varProp In this.BarProperties
            If varProp <> "Type" Then
                If dParent.Exists(varProp) Then
                    CallByName objItem, varProp, VbLet, dParent(varProp)
                    If Catch(438, 440) Then
                        ' No point logging those errors.
                    ElseIf CatchAny(eelNoError, vbNullString) Then
                        Debug.Print "Could not set " & varProp & " on " & objItem.Name
                    End If
                End If
            End If
        Next
        ' Now add any nested controls
        If dParent.Exists("Controls") Then
            For Each dControl In dParent("Controls")
                If dControl("BuiltIn") Then
                    ' We must first determine if it's an Access custom built-in control which requires special handling
                    Dim bolCustomBuiltIn As Boolean
                    Select Case dControl("Id")
                        Case 1835, 1836, 1837, 1838, 1839, 3885, 3886, 3888
                            bolCustomBuiltIn = True
                        Case 3887 ' We don't have diagrams yet. :-(
                            ' Throw warning
                            Log.Error eelWarning, "Custom commandbar opening database diagrams are not supported. If you want to support diagrams, please submit a PR including a working command bar with the diagram button.", ModuleName(Me) & ".BuildControls"
                        Case Else
                            bolCustomBuiltIn = False
                    End Select

                    If bolCustomBuiltIn Then
                        ' It's an Access "custom" built-in control which we cannot use .Add method.
                        ' We must instead copy from our template command bar into the target command bar.
                        ' So far, all custom built-in controls are buttons, so we'll use button rather than
                        ' more generic control.

                        ' Build the custom "built-in" control ourselves
                        BuildControls dControl, m_TemplateCommandBar.FindControl(ID:=dControl("Id")).Copy(objItem)
                    Else
                        ' Insert built-in control

                        ' A built-in control may fail to add. In this case, we need to warn the users
                        Dim objBuiltIn As CommandBarControl
                        Set objBuiltIn = objItem.Controls.Add(dControl("Type"), dControl("Id"))
                        If Not CatchAny(eelError, "Unable to add a built-in control using Id " & dControl("Id"), blnIncludeErrorWithDescription:=True) Then
                            objBuiltIn.Visible = dControl("Visible")
                        End If
                    End If
                Else
                    ' Build the control ourselves
                    BuildControls dControl, objItem.Controls.Add(dControl("Type"))
                End If
            Next dControl
        End If
    ElseIf TypeOf objItem Is CommandBarControl Then
        ' Set control properties
        For Each varProp In this.CtlWriteProperties
            If dParent.Exists(varProp) Then
                CallByName objItem, varProp, VbLet, dParent(varProp)
                If Catch(438, 440) Then
                    ' No point logging those errors
                ElseIf CatchAny(eelNoError, vbNullString) Then
                    Debug.Print "Could not set " & varProp & " on " & objItem.Caption
                End If
            End If
        Next varProp
        If TypeOf objItem Is CommandBarPopup And dParent.Exists("Controls") Then
            ' Add nested controls
            For Each dControl In dParent("Controls")
                BuildControls dControl, objItem.Controls.Add(dControl("Type"))
            Next dControl
        ElseIf TypeOf objItem Is CommandBarComboBox And dParent.Exists("List") Then
            Dim varItem As Variant
            Dim cbcItem As CommandBarComboBox
            Set objItem = cbcItem
            For Each varItem In dParent("List")
                cbcItem.AddItem varItem
            Next
        Else
            ' Check for picture
            If dParent.Exists("ImagePath") Then
                ' Build out full path to main image (Picture)
                strPath = GetPathFromRelative(dParent("ImagePath") & "Picture.bmp")
                If FSO.FileExists(strPath) Then
                    Set picItem = stdole.LoadPicture(strPath)
                    objItem.Picture = picItem
                End If
                ' Build path to image mask (transparency)
                strPath = GetPathFromRelative(dParent("ImagePath") & "Mask.bmp")
                If FSO.FileExists(strPath) Then
                    Set picItem = stdole.LoadPicture(strPath)
                    objItem.Mask = picItem
                End If
            End If
        End If
    Else
        ' Unsupported object type
        Stop
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetExisting
' Author    : Adam Waller
' Date      : 1/13/2024
' Purpose   : Returns any existing command bar by that name
'---------------------------------------------------------------------------------------
'
Private Function GetExisting(strName As String) As CommandBar
    LogUnhandledErrors
    On Error Resume Next
    Set GetExisting = CommandBars(strName)
    If Err Then Err.Clear
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetSource
' Author    : Adam Waller
' Date      : 2/14/2022
' Purpose   : Return the full content that will be saved to the source file.
'---------------------------------------------------------------------------------------
'
Private Function GetSource() As String
    GetSource = BuildJsonFile(TypeName(Me), GetDictionary, "CommandBar")
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetDictionary
' Author    : Adam Waller
' Date      : 5/17/2021
' Purpose   : Return a dictionary object of project properties.
'---------------------------------------------------------------------------------------
'
Private Function GetDictionary(Optional blnUseCache As Boolean = True) As Dictionary
    Dim dParent As Dictionary
    Set dParent = New Dictionary
    Set this.dImages = New Dictionary
    Set GetDictionary = BuildElementDictionary(dParent, m_CommandBar)
End Function


'---------------------------------------------------------------------------------------
' Procedure : BuildElementDictionary
' Author    : Adam Waller
' Date      : 1/12/2024
' Purpose   : A recursive function to build out the dictionary elements representing
'           : a CommandBar popup menu.
'           : objItem may represent a CommandBar, or a CommandBarControl.
'---------------------------------------------------------------------------------------
'
Private Function BuildElementDictionary(dParent As Dictionary, objItem As Object) As Dictionary

    Dim varProp As Variant
    Dim varCtl As Variant
    Dim varValue As Variant
    Dim varCtlProp As Variant
    Dim colItems As Collection
    Dim picItem As IPictureDisp
    Dim strPath As String

    LogUnhandledErrors
    On Error Resume Next

    If TypeOf objItem Is CommandBar Then
        ' Add command bar properties
        For Each varProp In this.BarProperties
            varValue = CallByName(objItem, varProp, VbGet)
            If Not CatchAny(eelNoError, vbNullString) Then
                dParent.Add varProp, varValue
            End If
        Next varProp
        ' Now add any nested controls
        Set colItems = New Collection
        dParent.Add "Controls", BuildControlCollection(objItem.Controls)
    ElseIf TypeOf objItem Is CommandBarControl Then
        ' Add control properties
        For Each varProp In this.CtlReadProperties
            varValue = Empty ' Clear the value from previous run
            varValue = CallByName(objItem, varProp, VbGet)
            ' No point in logging those 2 errors.
            If Catch(438, 440) Then
            ' Check for any other errors which should be logged just in case.
            ElseIf Not CatchAny(eelNoError, vbNullString) Then
                dParent.Add varProp, varValue
            End If
        Next varProp
        ' Save the ID for built-in items
        If objItem.BuiltIn Then
            dParent.Add "Id", objItem.ID
        End If
        If TypeOf objItem Is CommandBarComboBox Then
            If objItem.Type = msoControlDropdown Or objItem.Type = msoControlComboBox Then
            ' Build List
            Dim cbcItem As CommandBarComboBox
            Dim colList As Collection
            Dim i As Long
            Set cbcItem = objItem

            Set colList = New Collection
            For i = 1 To cbcItem.ListCount
                colList.Add cbcItem.List(i), CStr(i)
            Next
            dParent.Add "List", colList
            End If
        End If
        If TypeOf objItem Is CommandBarPopup Then
            ' Loop through nested controls
            dParent.Add "Controls", BuildControlCollection(objItem.Controls)
        Else
            ' Check for picture
            Set picItem = objItem.Picture
            ' Not all command bar controls have picture
            If Catch(438) Then
                ' Don't log this error and skip
            ElseIf CatchAny(eelNoError, vbNullString) Then
                ' Log any other error and skip
            Else
                ' Proceed with picture processing
                If Not picItem Is Nothing Then
                    strPath = GetImagePath(objItem)
                    ' Check path for possible duplicates (menu items with the same name)
                    If this.dImages.Exists(strPath) Then
                        ' Add a number to make the path unique.
                        strPath = strPath & "_" & this.dImages.Count
                    End If
                    ' Save reference to image objects to use when
                    ' exporting images to files.
                    this.dImages.Add strPath, objItem
                    ' Save path to images in element dictionary
                    dParent.Add "ImagePath", GetRelativePath(strPath)
                End If
            End If
        End If
    Else
        ' Unsupported object type
        Stop
    End If

    ' Return dictionary
    Set BuildElementDictionary = dParent

End Function


'---------------------------------------------------------------------------------------
' Procedure : BuildControlCollection
' Author    : Adam Waller
' Date      : 1/12/2024
' Purpose   : Return a collection of the command bar controls.
'---------------------------------------------------------------------------------------
'
Private Function BuildControlCollection(ctls As CommandBarControls) As Collection
    Dim ctl As CommandBarControl
    Dim dItem As Dictionary
    Set BuildControlCollection = New Collection
    For Each ctl In ctls
        Set dItem = New Dictionary
        BuildControlCollection.Add BuildElementDictionary(dItem, ctl)
    Next ctl
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetImagePath
' Author    : Adam Waller
' Date      : 1/12/2024
' Purpose   : Build a base path for the image files. (Image and mask)
'---------------------------------------------------------------------------------------
'
Private Function GetImagePath(ctl As CommandBarControl) As String

    Dim objParent As Object
    Dim ctlTest As Object
    Dim strName As String
    Dim strSegment As String

    On Error GoTo 0
    Set ctlTest = ctl
    ' Walk up the parent objects to build a full path to the control.
    ' (We are using a name path since the control IDs and indexes may change)
    Do
        If TypeOf ctlTest Is CommandBar Then
            ' Don't include command bar name, since we are using a subfolder
            ' to store images for each commmand bar.
            strSegment = vbNullString
        ElseIf TypeOf ctlTest Is CommandBarPopup Then
            ' We already have the name for this item
            strSegment = vbNullString
        Else
            ' Control item. (Use caption or ID)
            strSegment = Nz2(MultiReplace(ctlTest.Caption, _
                "&", vbNullString, _
                "...", vbNullString) _
                , ctlTest.ID)
        End If
        If Len(strSegment) Then strName = strSegment & "_" & strName
        If ctlTest.Parent Is Application Then
            ' This is the top level
            Exit Do
        Else
            ' Move up to parent object
            Set ctlTest = ctlTest.Parent
        End If
    Loop

    ' Convert name to a filesafe name
    GetImagePath = IDbComponent_BaseFolder & BuildPath2(GetSafeFileName(m_CommandBar.Name) & "_Images", GetSafeFileName(strName))

End Function


'---------------------------------------------------------------------------------------
' Procedure : Merge
' Author    : Adam Waller
' Date      : 5/17/2021
' Purpose   : Merge the source file into the existing database, updating or replacing
'           : any existing object.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Merge(strFile As String)
    ' Import if file exists
    If FSO.FileExists(strFile) Then
        IDbComponent_Import strFile
    Else
        VCSIndex.Remove Me, strFile
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : IDbComponent_MoveSource
' Author    : Adam Waller
' Date      : 9/10/2022
' Purpose   : Move the component's source file(s) from one folder to another
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_MoveSource(strFromFolder As String, strToFolder As String)
    MoveFileIfExists strFromFolder & FSO.GetFileName(IDbComponent_SourceFile), strToFolder
    MoveFolderIfExists strFromFolder & FSO.GetBaseName(IDbComponent_SourceFile), strToFolder
End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetAllFromDB
' Author    : Adam Waller
' Date      : 5/17/2021
' Purpose   : Return a collection of class objects represented by this component type.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_GetAllFromDB(Optional blnModifiedOnly As Boolean = False) As Dictionary

    Dim cCmdBar As IDbComponent
    Dim bar As CommandBar
    Dim blnAdd As Boolean

    ' Build collection if not already cached
    If m_Items(blnModifiedOnly) Is Nothing Then
        ' Loop through command bars, looking for any custom ones.

        ' TODO: Figure out a way to determine whether a command bar belongs to the current
        ' database. The command bars are associated with the application, not with the database
        ' so the application could have additional command bars loaded via add-in or other processes
        ' that are not actually a part of the database itself. The data is there in the internal
        ' table but not human readable.
        Set m_Items(blnModifiedOnly) = New Dictionary
        For Each bar In Application.CommandBars
            If bar.BuiltIn Then
                ' Skip
            ElseIf bar.Name = strTemplateCommandBarName Then
                ' Skip our template command bar
            Else
                Set cCmdBar = New clsDbCommandBar
                Set cCmdBar.DbObject = bar
                blnAdd = True
                If blnModifiedOnly Then blnAdd = cCmdBar.IsModified
                If blnAdd Then m_Items(blnModifiedOnly).Add cCmdBar.SourceFile, cCmdBar
            End If
        Next bar

    End If

    ' Return cached collection
    Set IDbComponent_GetAllFromDB = m_Items(blnModifiedOnly)

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetFileList
' Author    : Adam Waller
' Date      : 5/17/2021
' Purpose   : Return a list of file names to import for this component type.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_GetFileList() As Dictionary
    If m_FileList Is Nothing Then Set m_FileList = GetFilePathsInFolder(IDbComponent_BaseFolder, "*.json")
    Set IDbComponent_GetFileList = m_FileList
End Function


'---------------------------------------------------------------------------------------
' Procedure : IsModified
' Author    : Adam Waller
' Date      : 5/17/2021
' Purpose   : Returns true if the object in the database has been modified since
'           : the last export of the object.
'---------------------------------------------------------------------------------------
'
Public Function IDbComponent_IsModified() As Boolean
    IDbComponent_IsModified = VCSIndex.Item(Me).FileHash <> GetStringHash(GetSource, True)
End Function


'---------------------------------------------------------------------------------------
' Procedure : DateModified
' Author    : Adam Waller
' Date      : 5/17/2021
' Purpose   : The date/time the object was modified. (If possible to retrieve)
'           : If the modified date cannot be determined (such as application
'           : properties) then this function will return 0.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_DateModified() As Date
    IDbComponent_DateModified = 0
End Function


'---------------------------------------------------------------------------------------
' Procedure : Category
' Author    : Adam Waller
' Date      : 5/17/2021
' Purpose   : Return a category name for this type. (I.e. forms, queries, macros)
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_Category() As String
    IDbComponent_Category = "CommandBars"
End Property


'---------------------------------------------------------------------------------------
' Procedure : BaseFolder
' Author    : Adam Waller
' Date      : 5/17/2021
' Purpose   : Return the base folder for import/export of this component.
'---------------------------------------------------------------------------------------
Private Property Get IDbComponent_BaseFolder() As String
    IDbComponent_BaseFolder = Options.GetExportFolder & "menus" & PathSep
End Property


'---------------------------------------------------------------------------------------
' Procedure : FileExtensions
' Author    : Adam Waller
' Date      : 12/1/2023
' Purpose   : A collection of the file extensions used in source files for this
'           : component type.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_FileExtensions() As Collection
    Set IDbComponent_FileExtensions = New Collection
    IDbComponent_FileExtensions.Add "json"
End Property


'---------------------------------------------------------------------------------------
' Procedure : Name
' Author    : Adam Waller
' Date      : 5/17/2021
' Purpose   : Return a name to reference the object for use in logs and screen output.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_Name() As String
'    If m_Project Is Nothing Then Exit Property
    IDbComponent_Name = m_CommandBar.Name
End Property


'---------------------------------------------------------------------------------------
' Procedure : SourceFile
' Author    : Adam Waller
' Date      : 5/17/2021
' Purpose   : Return the full path of the source file for the current object.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_SourceFile() As String
    IDbComponent_SourceFile = IDbComponent_BaseFolder & GetSafeFileName(m_CommandBar.Name) & ".json"
End Property


'---------------------------------------------------------------------------------------
' Procedure : Count
' Author    : Adam Waller
' Date      : 5/17/2021
' Purpose   : Return a count of how many items are in this category.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_Count(Optional blnModifiedOnly As Boolean = False) As Long
    IDbComponent_Count = IDbComponent_GetAllFromDB(blnModifiedOnly).Count
End Property


'---------------------------------------------------------------------------------------
' Procedure : QuickCount
' Author    : Adam Waller
' Date      : 6/14/2022
' Purpose   : Return a cached, non-iterative approximate count of database objects
'           : for use with progress indicators when scanning for changes. Single file
'           : objects like database properties can simply return 1.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_QuickCount() As Long
    ' Probably 0 in most databases. We will have to iterate through the CommandBars
    ' to check the .BuiltIn flag to get an actual count.
    IDbComponent_QuickCount = 0
End Property


'---------------------------------------------------------------------------------------
' Procedure : ComponentType
' Author    : Adam Waller
' Date      : 5/17/2021
' Purpose   : The type of component represented by this class.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_ComponentType() As eDatabaseComponentType
    IDbComponent_ComponentType = edbCommandBar
End Property


'---------------------------------------------------------------------------------------
' Procedure : DbObject
' Author    : Adam Waller
' Date      : 5/17/2021
' Purpose   : This represents the database object we are dealing with.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_DbObject() As Object
    Set IDbComponent_DbObject = m_CommandBar
End Property
Private Property Set IDbComponent_DbObject(ByVal RHS As Object)
    Set m_CommandBar = RHS
End Property


'---------------------------------------------------------------------------------------
' Procedure : SingleFile
' Author    : Adam Waller
' Date      : 5/17/2021
' Purpose   : Returns true if the export of all items is done as a single file instead
'           : of individual files for each component. (I.e. properties, references)
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_SingleFile() As Boolean
    IDbComponent_SingleFile = False
End Property


'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : Adam Waller
' Date      : 5/17/2021
' Purpose   : Helps us know whether we have already counted the objects.
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
    'm_Count = -1

    ' We assume the add-in file already has the template command bar included. If it doesn't, then
    ' we will have problem importing Access' custom built-in buttons.
    Set m_TemplateCommandBar = Application.CommandBars(strTemplateCommandBarName)

    ' Set up property name lists

    ' CommandBar Properties
    Set this.BarProperties = New Collection
    With this.BarProperties
        .Add "AdaptiveMenu"
        .Add "Context"
        .Add "Enabled"
        .Add "Height"
        .Add "Left"
        .Add "NameLocal"
        .Add "Position"
        .Add "Protection"
        .Add "RowIndex"
        .Add "Top"
        .Add "Type"
        .Add "Visible"
        .Add "Width"
    End With

    ' CommmandBarControl Properties
    ' This should include all properties from different types of CommandBarControl
    ' even if they are not common among all types.

    ' Build a list for read-write properties to be used in building controls.
    Set this.CtlWriteProperties = New Collection
    With this.CtlWriteProperties
        .Add "BeginGroup"
        .Add "BuiltInFace"
        .Add "Caption"
        .Add "DescriptionText"
        .Add "DropDownLines"
        .Add "DropDownWidth"
        .Add "Enabled"
        .Add "FaceId"
        .Add "Height"
        .Add "HelpContextId"
        .Add "HelpFile"
        .Add "HyperlinkType"
        .Add "ListHeaderCount"
        .Add "ListIndex"
        .Add "OLEMenuGroup"
        .Add "OLEUsage"
        .Add "OnAction"
        .Add "Parameter"
        .Add "Priority"
        .Add "ShortcutText"
        .Add "State"
        .Add "Style"
        .Add "Tag"
        .Add "Text"
        .Add "TooltipText"
        .Add "Visible"
        .Add "Width"
    End With

    ' Build dictionary to be used for writing out to text files to include additional data about the
    ' controls so any changes can be detected via source code control
    Set this.CtlReadProperties = New Collection
    With this.CtlReadProperties
        .Add "BeginGroup"
        .Add "BuiltIn"
        .Add "BuiltInFace"
        .Add "Caption"
        .Add "DescriptionText"
        .Add "DropDownLines"
        .Add "DropDownWidth"
        .Add "Enabled"
        .Add "FaceId"
        .Add "Height"
        .Add "HelpContextId"
        .Add "HelpFile"
        .Add "HyperlinkType"
        .Add "Index"
        .Add "IsPriorityDropped"
        .Add "Left"
        .Add "ListCount"
        .Add "ListHeaderCount"
        .Add "ListIndex"
        .Add "OLEMenuGroup"
        .Add "OLEUsage"
        .Add "OnAction"
        .Add "Parameter"
        .Add "Priority"
        .Add "ShortcutText"
        .Add "State"
        .Add "Style"
        .Add "Tag"
        .Add "Text"
        .Add "TooltipText"
        .Add "Top"
        .Add "Type"
        .Add "Visible"
        .Add "Width"
    End With
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Class_Terminate
' Author    : Adam Waller
' Date      : 1/13/2024
' Purpose   : Clear any object references
'---------------------------------------------------------------------------------------
'
Private Sub Class_Terminate()
    ' Clear the images dictionary, just in case it still has references
    ' to the commandbar image objects.
    Set this.dImages = Nothing
End Sub


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
