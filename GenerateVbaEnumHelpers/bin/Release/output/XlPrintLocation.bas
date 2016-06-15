Attribute VB_Name = "wXlPrintLocation"
Function XlPrintLocationFromString(value As String) As XlPrintLocation
    If IsNumeric(value) Then
        XlPrintLocationFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlPrintSheetEnd": XlPrintLocationFromString = xlPrintSheetEnd
        Case "xlPrintInPlace": XlPrintLocationFromString = xlPrintInPlace
        Case "xlPrintNoComments": XlPrintLocationFromString = xlPrintNoComments
    End Select
End Function

Function XlPrintLocationToString(value As XlPrintLocation) As String
    Select Case value
        Case xlPrintSheetEnd: XlPrintLocationToString = "xlPrintSheetEnd"
        Case xlPrintInPlace: XlPrintLocationToString = "xlPrintInPlace"
        Case xlPrintNoComments: XlPrintLocationToString = "xlPrintNoComments"
    End Select
End Function
