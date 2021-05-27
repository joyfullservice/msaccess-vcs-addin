VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsConflictItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Category As String
Public FileName As String
Public ObjectDate As Date
Public IndexDate As Date
Public FileDate As Date
Public Resolution As eResolveConflict

