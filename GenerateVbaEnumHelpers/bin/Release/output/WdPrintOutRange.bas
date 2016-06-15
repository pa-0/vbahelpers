Attribute VB_Name = "wWdPrintOutRange"
Function WdPrintOutRangeFromString(value As String) As WdPrintOutRange
    If IsNumeric(value) Then
        WdPrintOutRangeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdPrintAllDocument": WdPrintOutRangeFromString = wdPrintAllDocument
        Case "wdPrintSelection": WdPrintOutRangeFromString = wdPrintSelection
        Case "wdPrintCurrentPage": WdPrintOutRangeFromString = wdPrintCurrentPage
        Case "wdPrintFromTo": WdPrintOutRangeFromString = wdPrintFromTo
        Case "wdPrintRangeOfPages": WdPrintOutRangeFromString = wdPrintRangeOfPages
    End Select
End Function

Function WdPrintOutRangeToString(value As WdPrintOutRange) As String
    Select Case value
        Case wdPrintAllDocument: WdPrintOutRangeToString = "wdPrintAllDocument"
        Case wdPrintSelection: WdPrintOutRangeToString = "wdPrintSelection"
        Case wdPrintCurrentPage: WdPrintOutRangeToString = "wdPrintCurrentPage"
        Case wdPrintFromTo: WdPrintOutRangeToString = "wdPrintFromTo"
        Case wdPrintRangeOfPages: WdPrintOutRangeToString = "wdPrintRangeOfPages"
    End Select
End Function
