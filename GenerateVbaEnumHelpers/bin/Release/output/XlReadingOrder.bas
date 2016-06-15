Attribute VB_Name = "wXlReadingOrder"
Function XlReadingOrderFromString(value As String) As XlReadingOrder
    If IsNumeric(value) Then
        XlReadingOrderFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlRTL": XlReadingOrderFromString = xlRTL
        Case "xlLTR": XlReadingOrderFromString = xlLTR
        Case "xlContext": XlReadingOrderFromString = xlContext
    End Select
End Function

Function XlReadingOrderToString(value As XlReadingOrder) As String
    Select Case value
        Case xlRTL: XlReadingOrderToString = "xlRTL"
        Case xlLTR: XlReadingOrderToString = "xlLTR"
        Case xlContext: XlReadingOrderToString = "xlContext"
    End Select
End Function
