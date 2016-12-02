Attribute VB_Name = "basTEMP"
Option Compare Database
Option Explicit



Public Sub InitializeSVN()

    ' VCS Settings for this database (Additional parameters may be added as needed)
    Dim varParams(0 To 3) As Variant
    varParams(0) = Array("System", "GitHub")    ' Set this first, before other settings.
    varParams(1) = Array("Export Folder", CodeDb.Name & ".src\")
    varParams(2) = Array("Show Debug", False)
    varParams(3) = Array("Include VBE", True)

    ' Clear current menu
    ReleaseObjectReferences
    
    LoadVersionControl varParams
    
End Sub
