Attribute VB_Name = "wXlEnableCancelKey"
Function XlEnableCancelKeyFromString(value As String) As XlEnableCancelKey
    If IsNumeric(value) Then
        XlEnableCancelKeyFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlDisabled": XlEnableCancelKeyFromString = xlDisabled
        Case "xlInterrupt": XlEnableCancelKeyFromString = xlInterrupt
        Case "xlErrorHandler": XlEnableCancelKeyFromString = xlErrorHandler
    End Select
End Function

Function XlEnableCancelKeyToString(value As XlEnableCancelKey) As String
    Select Case value
        Case xlDisabled: XlEnableCancelKeyToString = "xlDisabled"
        Case xlInterrupt: XlEnableCancelKeyToString = "xlInterrupt"
        Case xlErrorHandler: XlEnableCancelKeyToString = "xlErrorHandler"
    End Select
End Function
