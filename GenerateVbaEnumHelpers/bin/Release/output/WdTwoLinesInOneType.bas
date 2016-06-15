Attribute VB_Name = "wWdTwoLinesInOneType"
Function WdTwoLinesInOneTypeFromString(value As String) As WdTwoLinesInOneType
    If IsNumeric(value) Then
        WdTwoLinesInOneTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdTwoLinesInOneNone": WdTwoLinesInOneTypeFromString = wdTwoLinesInOneNone
        Case "wdTwoLinesInOneNoBrackets": WdTwoLinesInOneTypeFromString = wdTwoLinesInOneNoBrackets
        Case "wdTwoLinesInOneParentheses": WdTwoLinesInOneTypeFromString = wdTwoLinesInOneParentheses
        Case "wdTwoLinesInOneSquareBrackets": WdTwoLinesInOneTypeFromString = wdTwoLinesInOneSquareBrackets
        Case "wdTwoLinesInOneAngleBrackets": WdTwoLinesInOneTypeFromString = wdTwoLinesInOneAngleBrackets
        Case "wdTwoLinesInOneCurlyBrackets": WdTwoLinesInOneTypeFromString = wdTwoLinesInOneCurlyBrackets
    End Select
End Function

Function WdTwoLinesInOneTypeToString(value As WdTwoLinesInOneType) As String
    Select Case value
        Case wdTwoLinesInOneNone: WdTwoLinesInOneTypeToString = "wdTwoLinesInOneNone"
        Case wdTwoLinesInOneNoBrackets: WdTwoLinesInOneTypeToString = "wdTwoLinesInOneNoBrackets"
        Case wdTwoLinesInOneParentheses: WdTwoLinesInOneTypeToString = "wdTwoLinesInOneParentheses"
        Case wdTwoLinesInOneSquareBrackets: WdTwoLinesInOneTypeToString = "wdTwoLinesInOneSquareBrackets"
        Case wdTwoLinesInOneAngleBrackets: WdTwoLinesInOneTypeToString = "wdTwoLinesInOneAngleBrackets"
        Case wdTwoLinesInOneCurlyBrackets: WdTwoLinesInOneTypeToString = "wdTwoLinesInOneCurlyBrackets"
    End Select
End Function
