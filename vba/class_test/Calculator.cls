VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Calculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Basic VBA class module.
' Exercises: Class_Initialize, Public methods, Property Get/Let/Set.
Option Explicit

Private mValue As Double
Private mReady As Boolean
Private mTarget As Object

Private Sub Class_Initialize()
    mValue = 0
    mReady = True
End Sub

Public Property Get Value() As Double
    Value = mValue
End Property

Public Property Let Value(ByVal newValue As Double)
    mValue = newValue
End Property

Public Property Get Target() As Object
    Set Target = mTarget
End Property

Public Property Set Target(ByVal obj As Object)
    Set mTarget = obj
End Property

Public Function Add(ByVal a As Double, ByVal b As Double) As Double
    Add = a + b
End Function

Public Function Accumulate(ByVal n As Double) As Double
    mValue = mValue + n
    Accumulate = mValue
End Function

Public Function IsReady() As Boolean
    IsReady = mReady
End Function
