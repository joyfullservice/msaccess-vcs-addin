VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsJob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'@Folder("Infrastructure")

Public KeyHash As String
Public Action As String
Public Start As Double
Public Timeout As Double

' File-channel jobs keep their own result path so a reentrant launch cannot
' overwrite another job's handoff. Result survives queue removal because the
' caller holds this object. ResultExpired marks a timed-out or abandoned job
' so a late write is discarded instead of completing a later operation.
Public ResultFile As String
Public Result As Variant
Public ResultExpired As Boolean
