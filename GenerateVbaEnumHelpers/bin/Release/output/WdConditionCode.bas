Attribute VB_Name = "wWdConditionCode"
Function WdConditionCodeFromString(value As String) As WdConditionCode
    If IsNumeric(value) Then
        WdConditionCodeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdFirstRow": WdConditionCodeFromString = wdFirstRow
        Case "wdLastRow": WdConditionCodeFromString = wdLastRow
        Case "wdOddRowBanding": WdConditionCodeFromString = wdOddRowBanding
        Case "wdEvenRowBanding": WdConditionCodeFromString = wdEvenRowBanding
        Case "wdFirstColumn": WdConditionCodeFromString = wdFirstColumn
        Case "wdLastColumn": WdConditionCodeFromString = wdLastColumn
        Case "wdOddColumnBanding": WdConditionCodeFromString = wdOddColumnBanding
        Case "wdEvenColumnBanding": WdConditionCodeFromString = wdEvenColumnBanding
        Case "wdNECell": WdConditionCodeFromString = wdNECell
        Case "wdNWCell": WdConditionCodeFromString = wdNWCell
        Case "wdSECell": WdConditionCodeFromString = wdSECell
        Case "wdSWCell": WdConditionCodeFromString = wdSWCell
    End Select
End Function

Function WdConditionCodeToString(value As WdConditionCode) As String
    Select Case value
        Case wdFirstRow: WdConditionCodeToString = "wdFirstRow"
        Case wdLastRow: WdConditionCodeToString = "wdLastRow"
        Case wdOddRowBanding: WdConditionCodeToString = "wdOddRowBanding"
        Case wdEvenRowBanding: WdConditionCodeToString = "wdEvenRowBanding"
        Case wdFirstColumn: WdConditionCodeToString = "wdFirstColumn"
        Case wdLastColumn: WdConditionCodeToString = "wdLastColumn"
        Case wdOddColumnBanding: WdConditionCodeToString = "wdOddColumnBanding"
        Case wdEvenColumnBanding: WdConditionCodeToString = "wdEvenColumnBanding"
        Case wdNECell: WdConditionCodeToString = "wdNECell"
        Case wdNWCell: WdConditionCodeToString = "wdNWCell"
        Case wdSECell: WdConditionCodeToString = "wdSECell"
        Case wdSWCell: WdConditionCodeToString = "wdSWCell"
    End Select
End Function
