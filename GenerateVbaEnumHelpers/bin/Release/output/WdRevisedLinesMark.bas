Attribute VB_Name = "wWdRevisedLinesMark"
Function WdRevisedLinesMarkFromString(value As String) As WdRevisedLinesMark
    If IsNumeric(value) Then
        WdRevisedLinesMarkFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdRevisedLinesMarkNone": WdRevisedLinesMarkFromString = wdRevisedLinesMarkNone
        Case "wdRevisedLinesMarkLeftBorder": WdRevisedLinesMarkFromString = wdRevisedLinesMarkLeftBorder
        Case "wdRevisedLinesMarkRightBorder": WdRevisedLinesMarkFromString = wdRevisedLinesMarkRightBorder
        Case "wdRevisedLinesMarkOutsideBorder": WdRevisedLinesMarkFromString = wdRevisedLinesMarkOutsideBorder
    End Select
End Function

Function WdRevisedLinesMarkToString(value As WdRevisedLinesMark) As String
    Select Case value
        Case wdRevisedLinesMarkNone: WdRevisedLinesMarkToString = "wdRevisedLinesMarkNone"
        Case wdRevisedLinesMarkLeftBorder: WdRevisedLinesMarkToString = "wdRevisedLinesMarkLeftBorder"
        Case wdRevisedLinesMarkRightBorder: WdRevisedLinesMarkToString = "wdRevisedLinesMarkRightBorder"
        Case wdRevisedLinesMarkOutsideBorder: WdRevisedLinesMarkToString = "wdRevisedLinesMarkOutsideBorder"
    End Select
End Function
