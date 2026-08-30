Attribute VB_Name = "ClassBasicTests"
' Core class-module features: instantiation, methods, properties.
' Each function returns "PASS: ..." or "FAIL: ..." so the UNO runner can
' record a result per feature instead of dying on the first error.
Option Explicit

Public Function TestInstantiate() As String
    Dim c As Calculator
    Dim r As Double

    On Error Resume Next

    Set c = New Calculator
    If Err.Number <> 0 Then
        TestInstantiate = "FAIL: New raised " & Err.Number & " - " & Err.Description
        Exit Function
    End If

    If c Is Nothing Then
        TestInstantiate = "FAIL: instance is Nothing after New"
        Exit Function
    End If

    r = c.Add(2, 3)
    If Err.Number <> 0 Then
        TestInstantiate = "FAIL: Add raised " & Err.Number & " - " & Err.Description
        Exit Function
    End If

    If r = 5 Then
        TestInstantiate = "PASS: New Calculator + Add(2,3)=5"
    Else
        TestInstantiate = "FAIL: Add(2,3) returned " & r
    End If
End Function

Public Function TestDimAsNew() As String
    Dim c As New Calculator
    Dim r As Double

    On Error Resume Next

    r = c.Add(10, 7)
    If Err.Number <> 0 Then
        TestDimAsNew = "FAIL: Dim As New raised " & Err.Number & " - " & Err.Description
        Exit Function
    End If

    If r = 17 Then
        TestDimAsNew = "PASS: Dim As New auto-instantiation works"
    Else
        TestDimAsNew = "FAIL: Add(10,7) returned " & r
    End If
End Function

Public Function TestPropertyGetLet() As String
    Dim c As Calculator
    Dim v As Double

    On Error Resume Next

    Set c = New Calculator
    c.Value = 42
    If Err.Number <> 0 Then
        TestPropertyGetLet = "FAIL: Property Let raised " & Err.Number & " - " & Err.Description
        Exit Function
    End If

    v = c.Value
    If Err.Number <> 0 Then
        TestPropertyGetLet = "FAIL: Property Get raised " & Err.Number & " - " & Err.Description
        Exit Function
    End If

    If v = 42 Then
        TestPropertyGetLet = "PASS: Property Get/Let round-trips 42"
    Else
        TestPropertyGetLet = "FAIL: Property Get returned " & v
    End If
End Function

Public Function TestPropertySet() As String
    Dim c As Calculator
    Dim inner As Calculator
    Dim got As Object

    On Error Resume Next

    Set c = New Calculator
    Set inner = New Calculator
    inner.Value = 99

    Set c.Target = inner
    If Err.Number <> 0 Then
        TestPropertySet = "FAIL: Property Set raised " & Err.Number & " - " & Err.Description
        Exit Function
    End If

    Set got = c.Target
    If Err.Number <> 0 Or got Is Nothing Then
        TestPropertySet = "FAIL: Property Get (object) raised " & Err.Number & " - " & Err.Description
        Exit Function
    End If

    If got.Value = 99 Then
        TestPropertySet = "PASS: Property Set/Get holds an object reference"
    Else
        TestPropertySet = "FAIL: inner.Value came back as " & got.Value
    End If
End Function

Public Function TestPrivateState() As String
    Dim c As Calculator

    On Error Resume Next

    Set c = New Calculator
    c.Accumulate 5
    c.Accumulate 7
    If Err.Number <> 0 Then
        TestPrivateState = "FAIL: Accumulate raised " & Err.Number & " - " & Err.Description
        Exit Function
    End If

    If c.Value = 12 Then
        TestPrivateState = "PASS: private field persists across calls (5+7=12)"
    Else
        TestPrivateState = "FAIL: expected 12, got " & c.Value
    End If
End Function

Public Function TestClassInitialize() As String
    Dim c As Calculator

    On Error Resume Next

    Set c = New Calculator
    If Err.Number <> 0 Then
        TestClassInitialize = "FAIL: New raised " & Err.Number & " - " & Err.Description
        Exit Function
    End If

    If c.IsReady() = True Then
        TestClassInitialize = "PASS: Class_Initialize ran before first use"
    Else
        TestClassInitialize = "FAIL: Class_Initialize did not set the ready flag"
    End If
End Function

Public Function TestTypeName() As String
    Dim c As Calculator
    Dim n As String

    On Error Resume Next

    Set c = New Calculator
    n = TypeName(c)
    If Err.Number <> 0 Then
        TestTypeName = "FAIL: TypeName raised " & Err.Number & " - " & Err.Description
        Exit Function
    End If

    If n = "Calculator" Then
        TestTypeName = "PASS: TypeName() reports 'Calculator'"
    Else
        TestTypeName = "FAIL: TypeName() reported '" & n & "'"
    End If
End Function

Public Function TestCollectionOfObjects() As String
    Dim col As Collection
    Dim c As Calculator
    Dim i As Long
    Dim total As Double

    On Error Resume Next

    Set col = New Collection
    For i = 1 To 3
        Set c = New Calculator
        c.Value = i * 10
        col.Add c
    Next i

    If Err.Number <> 0 Then
        TestCollectionOfObjects = "FAIL: building collection raised " & Err.Number & " - " & Err.Description
        Exit Function
    End If

    For i = 1 To col.Count
        total = total + col.Item(i).Value
    Next i

    If Err.Number <> 0 Then
        TestCollectionOfObjects = "FAIL: reading collection raised " & Err.Number & " - " & Err.Description
        Exit Function
    End If

    If total = 60 Then
        TestCollectionOfObjects = "PASS: Collection holds class instances (10+20+30=60)"
    Else
        TestCollectionOfObjects = "FAIL: expected 60, got " & total
    End If
End Function
