Attribute VB_Name = "wXlCutCopyMode"
Function XlCutCopyModeFromString(value As String) As XlCutCopyMode
    If IsNumeric(value) Then
        XlCutCopyModeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlCopy": XlCutCopyModeFromString = xlCopy
        Case "xlCut": XlCutCopyModeFromString = xlCut
    End Select
End Function

Function XlCutCopyModeToString(value As XlCutCopyMode) As String
    Select Case value
        Case xlCopy: XlCutCopyModeToString = "xlCopy"
        Case xlCut: XlCutCopyModeToString = "xlCut"
    End Select
End Function
