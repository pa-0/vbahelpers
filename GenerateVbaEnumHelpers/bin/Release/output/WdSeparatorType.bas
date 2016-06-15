Attribute VB_Name = "wWdSeparatorType"
Function WdSeparatorTypeFromString(value As String) As WdSeparatorType
    If IsNumeric(value) Then
        WdSeparatorTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdSeparatorHyphen": WdSeparatorTypeFromString = wdSeparatorHyphen
        Case "wdSeparatorPeriod": WdSeparatorTypeFromString = wdSeparatorPeriod
        Case "wdSeparatorColon": WdSeparatorTypeFromString = wdSeparatorColon
        Case "wdSeparatorEmDash": WdSeparatorTypeFromString = wdSeparatorEmDash
        Case "wdSeparatorEnDash": WdSeparatorTypeFromString = wdSeparatorEnDash
    End Select
End Function

Function WdSeparatorTypeToString(value As WdSeparatorType) As String
    Select Case value
        Case wdSeparatorHyphen: WdSeparatorTypeToString = "wdSeparatorHyphen"
        Case wdSeparatorPeriod: WdSeparatorTypeToString = "wdSeparatorPeriod"
        Case wdSeparatorColon: WdSeparatorTypeToString = "wdSeparatorColon"
        Case wdSeparatorEmDash: WdSeparatorTypeToString = "wdSeparatorEmDash"
        Case wdSeparatorEnDash: WdSeparatorTypeToString = "wdSeparatorEnDash"
    End Select
End Function
