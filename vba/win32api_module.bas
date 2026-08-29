Attribute VB_Name = "Win32ApiModule"
' Win32 API Test Module
' This module demonstrates calling Windows API functions from VBA
' Tests whether LibreOffice VBA compatibility supports Win32 API calls

Option Explicit

' ============================================
' Win32 API Declarations
' ============================================

#If VBA7 Then
    ' 64-bit Office declarations
    Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
    Private Declare PtrSafe Function GetComputerNameA Lib "kernel32" _
        (ByVal lpBuffer As String, ByRef nSize As Long) As Long
    Private Declare PtrSafe Function GetUserNameA Lib "advapi32.dll" _
        (ByVal lpBuffer As String, ByRef nSize As Long) As Long
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Private Declare PtrSafe Function GetSystemMetrics Lib "user32" _
        (ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function GetTempPathA Lib "kernel32" _
        (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
#Else
    ' 32-bit Office declarations
    Private Declare Function GetTickCount Lib "kernel32" () As Long
    Private Declare Function GetComputerNameA Lib "kernel32" _
        (ByVal lpBuffer As String, nSize As Long) As Long
    Private Declare Function GetUserNameA Lib "advapi32.dll" _
        (ByVal lpBuffer As String, nSize As Long) As Long
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Private Declare Function GetSystemMetrics Lib "user32" _
        (ByVal nIndex As Long) As Long
    Private Declare Function GetTempPathA Lib "kernel32" _
        (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
#End If

' System Metrics constants
Private Const SM_CXSCREEN As Long = 0
Private Const SM_CYSCREEN As Long = 1

' ============================================
' Test Functions
' ============================================

' Get the computer name using Win32 API
Public Function GetComputerNameAPI() As String
    Dim buffer As String
    Dim bufferSize As Long

    buffer = String$(256, 0)
    bufferSize = Len(buffer)

    If GetComputerNameA(buffer, bufferSize) <> 0 Then
        GetComputerNameAPI = Left$(buffer, bufferSize)
    Else
        GetComputerNameAPI = "Unknown"
    End If
End Function

' Get the current user name using Win32 API
Public Function GetUserNameAPI() As String
    Dim buffer As String
    Dim bufferSize As Long

    buffer = String$(256, 0)
    bufferSize = Len(buffer)

    If GetUserNameA(buffer, bufferSize) <> 0 Then
        GetUserNameAPI = Left$(buffer, bufferSize - 1)
    Else
        GetUserNameAPI = "Unknown"
    End If
End Function

' Get system uptime in milliseconds
Public Function GetSystemUptime() As Long
    GetSystemUptime = GetTickCount()
End Function

' Get screen resolution
Public Function GetScreenWidth() As Long
    GetScreenWidth = GetSystemMetrics(SM_CXSCREEN)
End Function

Public Function GetScreenHeight() As Long
    GetScreenHeight = GetSystemMetrics(SM_CYSCREEN)
End Function

' Get temp path
Public Function GetTempPath() As String
    Dim buffer As String
    Dim length As Long

    buffer = String$(260, 0)
    length = GetTempPathA(260, buffer)

    If length > 0 Then
        GetTempPath = Left$(buffer, length)
    Else
        GetTempPath = ""
    End If
End Function

' Sleep for specified milliseconds
Public Sub SleepMS(ByVal milliseconds As Long)
    Sleep milliseconds
End Sub

' ============================================
' Main Test Runner
' ============================================

' Run all Win32 API tests and write results to worksheet
Public Sub RunWin32ApiTests()
    Dim ws As Object
    Dim row As Long
    Dim startTick As Long
    Dim endTick As Long

    On Error Resume Next

    Set ws = ThisWorkbook.Sheets(1)
    row = 1

    ' Header
    ws.Cells(row, 1).Value = "Win32 API Test Results"
    ws.Cells(row, 2).Value = Now()
    row = row + 2

    ' Test 1: GetComputerName
    ws.Cells(row, 1).Value = "Computer Name:"
    ws.Cells(row, 2).Value = GetComputerNameAPI()
    If Err.Number <> 0 Then
        ws.Cells(row, 2).Value = "ERROR: " & Err.Description
        Err.Clear
    End If
    row = row + 1

    ' Test 2: GetUserName
    ws.Cells(row, 1).Value = "User Name:"
    ws.Cells(row, 2).Value = GetUserNameAPI()
    If Err.Number <> 0 Then
        ws.Cells(row, 2).Value = "ERROR: " & Err.Description
        Err.Clear
    End If
    row = row + 1

    ' Test 3: GetTickCount
    ws.Cells(row, 1).Value = "System Uptime (ms):"
    ws.Cells(row, 2).Value = GetSystemUptime()
    If Err.Number <> 0 Then
        ws.Cells(row, 2).Value = "ERROR: " & Err.Description
        Err.Clear
    End If
    row = row + 1

    ' Test 4: Screen Resolution
    ws.Cells(row, 1).Value = "Screen Width:"
    ws.Cells(row, 2).Value = GetScreenWidth()
    If Err.Number <> 0 Then
        ws.Cells(row, 2).Value = "ERROR: " & Err.Description
        Err.Clear
    End If
    row = row + 1

    ws.Cells(row, 1).Value = "Screen Height:"
    ws.Cells(row, 2).Value = GetScreenHeight()
    If Err.Number <> 0 Then
        ws.Cells(row, 2).Value = "ERROR: " & Err.Description
        Err.Clear
    End If
    row = row + 1

    ' Test 5: Temp Path
    ws.Cells(row, 1).Value = "Temp Path:"
    ws.Cells(row, 2).Value = GetTempPath()
    If Err.Number <> 0 Then
        ws.Cells(row, 2).Value = "ERROR: " & Err.Description
        Err.Clear
    End If
    row = row + 1

    ' Test 6: Sleep timing test
    ws.Cells(row, 1).Value = "Sleep Test (100ms):"
    startTick = GetTickCount()
    SleepMS 100
    endTick = GetTickCount()
    ws.Cells(row, 2).Value = "Elapsed: " & (endTick - startTick) & " ms"
    If Err.Number <> 0 Then
        ws.Cells(row, 2).Value = "ERROR: " & Err.Description
        Err.Clear
    End If
    row = row + 1

    ' Summary
    row = row + 1
    ws.Cells(row, 1).Value = "Test Complete"

    ' Save workbook
    ThisWorkbook.Save

    On Error GoTo 0
End Sub
