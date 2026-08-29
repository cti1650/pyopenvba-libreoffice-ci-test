Attribute VB_Name = "SampleModule"
' Sample VBA Module for testing pyOpenVBA and LibreOffice
' This module contains simple functions that can be tested in CI

Option Explicit

' Simple function to add two numbers
Public Function AddNumbers(a As Double, b As Double) As Double
    AddNumbers = a + b
End Function

' Simple function to concatenate strings
Public Function ConcatStrings(str1 As String, str2 As String) As String
    ConcatStrings = str1 & str2
End Function

' Function to write result to a cell (for testing)
Public Sub WriteTestResult()
    Dim ws As Object
    Set ws = ThisWorkbook.Sheets(1)
    ws.Range("A1").Value = "VBA Test"
    ws.Range("B1").Value = AddNumbers(10, 20)
    ws.Range("C1").Value = ConcatStrings("Hello", "World")
End Sub

' Main test entry point
Public Sub RunAllTests()
    Dim result As String
    result = ""

    ' Test AddNumbers
    If AddNumbers(2, 3) = 5 Then
        result = result & "AddNumbers: PASS" & vbCrLf
    Else
        result = result & "AddNumbers: FAIL" & vbCrLf
    End If

    ' Test ConcatStrings
    If ConcatStrings("Hello", "World") = "HelloWorld" Then
        result = result & "ConcatStrings: PASS" & vbCrLf
    Else
        result = result & "ConcatStrings: FAIL" & vbCrLf
    End If

    ' Output results to cell
    ThisWorkbook.Sheets(1).Range("A1").Value = result
    ThisWorkbook.Save
End Sub
