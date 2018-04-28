Attribute VB_Name = "VCS_Button_Functions"
Option Compare Database

' function to call update functions from button clicks
Function subUpdateBtn(btnFunction As String)
    ' do this every time
    ' loadVCS

    ' update form not tables
    ' export form
    ' reset form, w/ lookup tables, prompt user to confirm
    '& btnFunction
    Select Case btnFunction
        Case "updateFormsBtn" ', "exportFormsBtn", "resetFormsBtn"
            Debug.Print "button worked: " & btnFunction
            ImportProject (True)
        Case "exportFormsBtn"
            Debug.Print "button worked2: " & btnFunction
            ExportAllSource (True) ' will skip exporting tables
        Case Else
            MsgBox "current function doesn't yet exist"
    End Select

End Function
' filler for things like reset
Function formDialog()
    
    'Variable Declaration
    Dim OutPut As Integer

    'Example of vbDefaultButton2
    OutPut = MsgBox("Close the File.Try Again?", vbYesNoCancel + vbDefaultButton3, "Example of vbDefaultButton3")

End Function
