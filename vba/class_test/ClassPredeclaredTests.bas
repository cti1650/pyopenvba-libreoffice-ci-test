Attribute VB_Name = "ClassPredeclaredTests"
' VB_PredeclaredId = True: the class should be usable without New.
Option Explicit

Public Function TestPredeclaredInstance() As String
    Dim n As String

    On Error Resume Next

    n = AppConfig.AppName
    If Err.Number <> 0 Then
        TestPredeclaredInstance = "FAIL: default instance raised " & Err.Number & " - " & Err.Description
        Exit Function
    End If

    If n = "PredeclaredDefault" Then
        TestPredeclaredInstance = "PASS: default instance usable without New"
    Else
        TestPredeclaredInstance = "FAIL: AppName returned '" & n & "'"
    End If
End Function

Public Function TestPredeclaredIsShared() As String
    On Error Resume Next

    AppConfig.AppName = "Mutated"
    If Err.Number <> 0 Then
        TestPredeclaredIsShared = "FAIL: writing default instance raised " & Err.Number & " - " & Err.Description
        Exit Function
    End If

    If AppConfig.AppName = "Mutated" Then
        TestPredeclaredIsShared = "PASS: default instance keeps state across statements"
    Else
        TestPredeclaredIsShared = "FAIL: AppName came back as '" & AppConfig.AppName & "'"
    End If
End Function
