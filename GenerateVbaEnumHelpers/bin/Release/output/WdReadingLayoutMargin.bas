Attribute VB_Name = "wWdReadingLayoutMargin"
Function WdReadingLayoutMarginFromString(value As String) As WdReadingLayoutMargin
    If IsNumeric(value) Then
        WdReadingLayoutMarginFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdAutomaticMargin": WdReadingLayoutMarginFromString = wdAutomaticMargin
        Case "wdSuppressMargin": WdReadingLayoutMarginFromString = wdSuppressMargin
        Case "wdFullMargin": WdReadingLayoutMarginFromString = wdFullMargin
    End Select
End Function

Function WdReadingLayoutMarginToString(value As WdReadingLayoutMargin) As String
    Select Case value
        Case wdAutomaticMargin: WdReadingLayoutMarginToString = "wdAutomaticMargin"
        Case wdSuppressMargin: WdReadingLayoutMarginToString = "wdSuppressMargin"
        Case wdFullMargin: WdReadingLayoutMarginToString = "wdFullMargin"
    End Select
End Function
