Attribute VB_Name = "wXlRobustConnect"
Function XlRobustConnectFromString(value As String) As XlRobustConnect
    If IsNumeric(value) Then
        XlRobustConnectFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlAsRequired": XlRobustConnectFromString = xlAsRequired
        Case "xlAlways": XlRobustConnectFromString = xlAlways
        Case "xlNever": XlRobustConnectFromString = xlNever
    End Select
End Function

Function XlRobustConnectToString(value As XlRobustConnect) As String
    Select Case value
        Case xlAsRequired: XlRobustConnectToString = "xlAsRequired"
        Case xlAlways: XlRobustConnectToString = "xlAlways"
        Case xlNever: XlRobustConnectToString = "xlNever"
    End Select
End Function
