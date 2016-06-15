Attribute VB_Name = "wPpPrintHandoutOrder"
Function PpPrintHandoutOrderFromString(value As String) As PpPrintHandoutOrder
    If IsNumeric(value) Then
        PpPrintHandoutOrderFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "ppPrintHandoutVerticalFirst": PpPrintHandoutOrderFromString = ppPrintHandoutVerticalFirst
        Case "ppPrintHandoutHorizontalFirst": PpPrintHandoutOrderFromString = ppPrintHandoutHorizontalFirst
    End Select
End Function

Function PpPrintHandoutOrderToString(value As PpPrintHandoutOrder) As String
    Select Case value
        Case ppPrintHandoutVerticalFirst: PpPrintHandoutOrderToString = "ppPrintHandoutVerticalFirst"
        Case ppPrintHandoutHorizontalFirst: PpPrintHandoutOrderToString = "ppPrintHandoutHorizontalFirst"
    End Select
End Function
