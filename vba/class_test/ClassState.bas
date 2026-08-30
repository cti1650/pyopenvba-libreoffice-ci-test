Attribute VB_Name = "ClassState"
' Shared counters written by Tracked's lifecycle handlers.
Option Explicit

Public InitCount As Long
Public TermCount As Long

Public Sub ResetCounters()
    InitCount = 0
    TermCount = 0
End Sub
