Attribute VB_Name = "modProc"
Option Explicit
Public Function TokenByPos(stg As String, intIdx As Integer, stgSep As String) As String

    ' intIdx = -1 to get last element in the string
    
    Dim stgSplit
    
    If Right$(stg, 1) <> stgSep Then
        stgSplit = Split(stg & stgSep, stgSep, -1, 1)
    Else
        stgSplit = Split(stg, stgSep, -1, 1)
    End If
    If intIdx < 0 Then
        TokenByPos = stgSplit(UBound(stgSplit) - 1)
    ElseIf intIdx <= (UBound(stgSplit)) Then
        TokenByPos = stgSplit(intIdx - 1)
    Else
        TokenByPos = Empty
    End If

End Function


