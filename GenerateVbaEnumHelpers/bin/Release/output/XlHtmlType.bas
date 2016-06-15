Attribute VB_Name = "wXlHtmlType"
Function XlHtmlTypeFromString(value As String) As XlHtmlType
    If IsNumeric(value) Then
        XlHtmlTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlHtmlStatic": XlHtmlTypeFromString = xlHtmlStatic
        Case "xlHtmlCalc": XlHtmlTypeFromString = xlHtmlCalc
        Case "xlHtmlList": XlHtmlTypeFromString = xlHtmlList
        Case "xlHtmlChart": XlHtmlTypeFromString = xlHtmlChart
    End Select
End Function

Function XlHtmlTypeToString(value As XlHtmlType) As String
    Select Case value
        Case xlHtmlStatic: XlHtmlTypeToString = "xlHtmlStatic"
        Case xlHtmlCalc: XlHtmlTypeToString = "xlHtmlCalc"
        Case xlHtmlList: XlHtmlTypeToString = "xlHtmlList"
        Case xlHtmlChart: XlHtmlTypeToString = "xlHtmlChart"
    End Select
End Function
