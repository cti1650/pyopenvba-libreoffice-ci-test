VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventListener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Sinks EventSource.ValueChanged via WithEvents.
Option Explicit

Private WithEvents mSource As EventSource

Public LastReceived As Long
Public EventCount As Long

Public Sub Attach(ByVal src As EventSource)
    Set mSource = src
End Sub

Private Sub mSource_ValueChanged(ByVal newValue As Long)
    LastReceived = newValue
    EventCount = EventCount + 1
End Sub
