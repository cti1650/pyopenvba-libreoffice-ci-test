VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Circle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Implements IShape. LibreOffice Basic's VBA mode is not expected to
' support Implements; this class exists to measure exactly how it fails.
Option Explicit

Implements IShape

Private mRadius As Double

Public Property Let Radius(ByVal r As Double)
    mRadius = r
End Property

Private Function IShape_Area() As Double
    IShape_Area = 3.14159265358979 * mRadius * mRadius
End Function

Private Property Get IShape_ShapeName() As String
    IShape_ShapeName = "Circle"
End Property
