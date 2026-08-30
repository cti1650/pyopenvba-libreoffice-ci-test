Attribute VB_Name = "ClassEventsTests"
' Public Event / RaiseEvent / WithEvents. Isolated in its own module so a
' compile failure here does not take the other test modules down with it.
Option Explicit

Public Function TestEventClassesCompile() As String
    Dim src As EventSource
    Dim lis As EventListener

    On Error Resume Next

    Set src = New EventSource
    If Err.Number <> 0 Then
        TestEventClassesCompile = "FAIL: New EventSource raised " & Err.Number & " - " & Err.Description
        Exit Function
    End If

    Set lis = New EventListener
    If Err.Number <> 0 Then
        TestEventClassesCompile = "FAIL: New EventListener (WithEvents) raised " & Err.Number & " - " & Err.Description
        Exit Function
    End If

    TestEventClassesCompile = "PASS: event source and WithEvents listener both instantiate"
End Function

Public Function TestRaiseEventDelivered() As String
    Dim src As EventSource
    Dim lis As EventListener

    On Error Resume Next

    Set src = New EventSource
    Set lis = New EventListener

    lis.Attach src
    If Err.Number <> 0 Then
        TestRaiseEventDelivered = "FAIL: Attach raised " & Err.Number & " - " & Err.Description
        Exit Function
    End If

    src.SetValue 7
    If Err.Number <> 0 Then
        TestRaiseEventDelivered = "FAIL: SetValue/RaiseEvent raised " & Err.Number & " - " & Err.Description
        Exit Function
    End If

    If lis.EventCount = 1 And lis.LastReceived = 7 Then
        TestRaiseEventDelivered = "PASS: RaiseEvent delivered to WithEvents handler"
    Else
        TestRaiseEventDelivered = "FAIL: EventCount=" & lis.EventCount & " LastReceived=" & lis.LastReceived
    End If
End Function
