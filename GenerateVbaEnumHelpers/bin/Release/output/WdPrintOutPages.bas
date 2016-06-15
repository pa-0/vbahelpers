Attribute VB_Name = "wWdPrintOutPages"
Function WdPrintOutPagesFromString(value As String) As WdPrintOutPages
    If IsNumeric(value) Then
        WdPrintOutPagesFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdPrintAllPages": WdPrintOutPagesFromString = wdPrintAllPages
        Case "wdPrintOddPagesOnly": WdPrintOutPagesFromString = wdPrintOddPagesOnly
        Case "wdPrintEvenPagesOnly": WdPrintOutPagesFromString = wdPrintEvenPagesOnly
    End Select
End Function

Function WdPrintOutPagesToString(value As WdPrintOutPages) As String
    Select Case value
        Case wdPrintAllPages: WdPrintOutPagesToString = "wdPrintAllPages"
        Case wdPrintOddPagesOnly: WdPrintOutPagesToString = "wdPrintOddPagesOnly"
        Case wdPrintEvenPagesOnly: WdPrintOutPagesToString = "wdPrintEvenPagesOnly"
    End Select
End Function
