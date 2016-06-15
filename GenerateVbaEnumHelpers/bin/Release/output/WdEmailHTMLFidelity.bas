Attribute VB_Name = "wWdEmailHTMLFidelity"
Function WdEmailHTMLFidelityFromString(value As String) As WdEmailHTMLFidelity
    If IsNumeric(value) Then
        WdEmailHTMLFidelityFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdEmailHTMLFidelityLow": WdEmailHTMLFidelityFromString = wdEmailHTMLFidelityLow
        Case "wdEmailHTMLFidelityMedium": WdEmailHTMLFidelityFromString = wdEmailHTMLFidelityMedium
        Case "wdEmailHTMLFidelityHigh": WdEmailHTMLFidelityFromString = wdEmailHTMLFidelityHigh
    End Select
End Function

Function WdEmailHTMLFidelityToString(value As WdEmailHTMLFidelity) As String
    Select Case value
        Case wdEmailHTMLFidelityLow: WdEmailHTMLFidelityToString = "wdEmailHTMLFidelityLow"
        Case wdEmailHTMLFidelityMedium: WdEmailHTMLFidelityToString = "wdEmailHTMLFidelityMedium"
        Case wdEmailHTMLFidelityHigh: WdEmailHTMLFidelityToString = "wdEmailHTMLFidelityHigh"
    End Select
End Function
