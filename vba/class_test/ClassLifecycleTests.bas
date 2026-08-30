Attribute VB_Name = "ClassLifecycleTests"
' Class_Initialize / Class_Terminate observation via ClassState counters.
Option Explicit

Public Function TestInitializeCounter() As String
    Dim t As Tracked

    On Error Resume Next

    ResetCounters
    Set t = New Tracked
    If Err.Number <> 0 Then
        TestInitializeCounter = "FAIL: New Tracked raised " & Err.Number & " - " & Err.Description
        Exit Function
    End If

    ' Touch the instance so an optimizing runtime cannot skip construction.
    If t.Ping() <> "pong" Then
        TestInitializeCounter = "FAIL: Ping did not return 'pong'"
        Exit Function
    End If

    If InitCount = 1 Then
        TestInitializeCounter = "PASS: Class_Initialize fired exactly once"
    Else
        TestInitializeCounter = "FAIL: InitCount = " & InitCount & " (expected 1)"
    End If
End Function

Public Function TestTerminateCounter() As String
    Dim t As Tracked

    On Error Resume Next

    ResetCounters
    Set t = New Tracked
    If t.Ping() <> "pong" Then
        TestTerminateCounter = "FAIL: Ping did not return 'pong'"
        Exit Function
    End If

    Set t = Nothing
    If Err.Number <> 0 Then
        TestTerminateCounter = "FAIL: Set Nothing raised " & Err.Number & " - " & Err.Description
        Exit Function
    End If

    If TermCount = 1 Then
        TestTerminateCounter = "PASS: Class_Terminate fired on Set Nothing"
    Else
        TestTerminateCounter = "FAIL: TermCount = " & TermCount & " (expected 1)"
    End If
End Function
