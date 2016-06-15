Attribute VB_Name = "wWdReadingOrder"
Function WdReadingOrderFromString(value As String) As WdReadingOrder
    If IsNumeric(value) Then
        WdReadingOrderFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdReadingOrderRtl": WdReadingOrderFromString = wdReadingOrderRtl
        Case "wdReadingOrderLtr": WdReadingOrderFromString = wdReadingOrderLtr
    End Select
End Function

Function WdReadingOrderToString(value As WdReadingOrder) As String
    Select Case value
        Case wdReadingOrderRtl: WdReadingOrderToString = "wdReadingOrderRtl"
        Case wdReadingOrderLtr: WdReadingOrderToString = "wdReadingOrderLtr"
    End Select
End Function
