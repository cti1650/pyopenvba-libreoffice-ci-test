VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tracked"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Lifecycle probe: bumps counters in ClassState so the caller can observe
' whether LibreOffice fires Class_Initialize / Class_Terminate at all.
Option Explicit

Private Sub Class_Initialize()
    InitCount = InitCount + 1
End Sub

Private Sub Class_Terminate()
    TermCount = TermCount + 1
End Sub

Public Function Ping() As String
    Ping = "pong"
End Function
