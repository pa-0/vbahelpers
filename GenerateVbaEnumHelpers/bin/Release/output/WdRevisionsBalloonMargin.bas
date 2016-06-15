Attribute VB_Name = "wWdRevisionsBalloonMargin"
Function WdRevisionsBalloonMarginFromString(value As String) As WdRevisionsBalloonMargin
    If IsNumeric(value) Then
        WdRevisionsBalloonMarginFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdLeftMargin": WdRevisionsBalloonMarginFromString = wdLeftMargin
        Case "wdRightMargin": WdRevisionsBalloonMarginFromString = wdRightMargin
    End Select
End Function

Function WdRevisionsBalloonMarginToString(value As WdRevisionsBalloonMargin) As String
    Select Case value
        Case wdLeftMargin: WdRevisionsBalloonMarginToString = "wdLeftMargin"
        Case wdRightMargin: WdRevisionsBalloonMarginToString = "wdRightMargin"
    End Select
End Function
