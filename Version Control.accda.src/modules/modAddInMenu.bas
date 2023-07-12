Attribute VB_Name = "modAddInMenu"
'---------------------------------------------------------------------------------------
' Module    : modAddIn
' Author    : Adam Waller
' Date      : 12/4/2020
' Purpose   : Functions called by the Database Tools -> Add-Ins menu.
'           : (This was the original way to invoke the add-in before the COM ribbon
'           :  project provided a full ribbon menu.)
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit


'---------------------------------------------------------------------------------------
' Procedure : AddInMenuItemLaunch
' Author    : Adam Waller
' Date      : 1/14/2020
' Purpose   : Launch the main add-in form.
'---------------------------------------------------------------------------------------
'
Public Function AddInMenuItemLaunch()
    VCS.Show
End Function


'---------------------------------------------------------------------------------------
' Procedure : AddInOptionsLaunch
' Author    : Hecon5
' Date      : 2/05/2020
' Purpose   : Launch the main add-in form.
'---------------------------------------------------------------------------------------
'
Public Function AddInOptionsLaunch()
    VCS.ShowOptions
End Function


'---------------------------------------------------------------------------------------
' Procedure : AddInMenuItemExport
' Author    : Adam Waller
' Date      : 4/15/2020
' Purpose   : Open main form and start export immediately. (Save users a click)
'---------------------------------------------------------------------------------------
'
Public Function AddInMenuItemExport()
    VCS.Export
End Function
