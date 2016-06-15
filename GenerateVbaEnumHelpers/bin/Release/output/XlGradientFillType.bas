Attribute VB_Name = "wXlGradientFillType"
Function XlGradientFillTypeFromString(value As String) As XlGradientFillType
    If IsNumeric(value) Then
        XlGradientFillTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlGradientFillLinear": XlGradientFillTypeFromString = xlGradientFillLinear
        Case "xlGradientFillPath": XlGradientFillTypeFromString = xlGradientFillPath
    End Select
End Function

Function XlGradientFillTypeToString(value As XlGradientFillType) As String
    Select Case value
        Case xlGradientFillLinear: XlGradientFillTypeToString = "xlGradientFillLinear"
        Case xlGradientFillPath: XlGradientFillTypeToString = "xlGradientFillPath"
    End Select
End Function
