VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Raises a custom event. Paired with EventListener (WithEvents).
Option Explicit

Public Event ValueChanged(ByVal newValue As Long)

Private mValue As Long

Public Sub SetValue(ByVal v As Long)
    mValue = v
    RaiseEvent ValueChanged(v)
End Sub

Public Property Get Value() As Long
    Value = mValue
End Property
