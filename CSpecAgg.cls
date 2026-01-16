VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSpecAgg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public spec As String
Public qty As Long
Public totalW As Double
Public details As Collection

Private Sub Class_Initialize()
    Set details = New Collection
End Sub


