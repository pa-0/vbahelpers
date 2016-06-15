Attribute VB_Name = "wXlWindowType"
Function XlWindowTypeFromString(value As String) As XlWindowType
    If IsNumeric(value) Then
        XlWindowTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlWorkbook": XlWindowTypeFromString = xlWorkbook
        Case "xlClipboard": XlWindowTypeFromString = xlClipboard
        Case "xlChartInPlace": XlWindowTypeFromString = xlChartInPlace
        Case "xlChartAsWindow": XlWindowTypeFromString = xlChartAsWindow
        Case "xlInfo": XlWindowTypeFromString = xlInfo
    End Select
End Function

Function XlWindowTypeToString(value As XlWindowType) As String
    Select Case value
        Case xlWorkbook: XlWindowTypeToString = "xlWorkbook"
        Case xlClipboard: XlWindowTypeToString = "xlClipboard"
        Case xlChartInPlace: XlWindowTypeToString = "xlChartInPlace"
        Case xlChartAsWindow: XlWindowTypeToString = "xlChartAsWindow"
        Case xlInfo: XlWindowTypeToString = "xlInfo"
    End Select
End Function
