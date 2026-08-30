Attribute VB_Name = "ClassImplementsTests"
' Implements / interface polymorphism. Isolated in its own module so a
' compile failure here does not take the other test modules down with it.
Option Explicit

Public Function TestImplementsCompiles() As String
    Dim c As Circle

    On Error Resume Next

    Set c = New Circle
    If Err.Number <> 0 Then
        TestImplementsCompiles = "FAIL: New Circle raised " & Err.Number & " - " & Err.Description
        Exit Function
    End If

    TestImplementsCompiles = "PASS: class using Implements instantiates"
End Function

Public Function TestInterfacePolymorphism() As String
    Dim c As Circle
    Dim s As IShape
    Dim a As Double

    On Error Resume Next

    Set c = New Circle
    c.Radius = 2

    Set s = c
    If Err.Number <> 0 Then
        TestInterfacePolymorphism = "FAIL: assign to IShape raised " & Err.Number & " - " & Err.Description
        Exit Function
    End If

    a = s.Area()
    If Err.Number <> 0 Then
        TestInterfacePolymorphism = "FAIL: IShape.Area raised " & Err.Number & " - " & Err.Description
        Exit Function
    End If

    If Abs(a - 12.5663706143592) < 0.000001 Then
        TestInterfacePolymorphism = "PASS: interface dispatch returns Area=" & a
    Else
        TestInterfacePolymorphism = "FAIL: Area returned " & a
    End If
End Function
