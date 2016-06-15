Attribute VB_Name = "wWdShowFilter"
Function WdShowFilterFromString(value As String) As WdShowFilter
    If IsNumeric(value) Then
        WdShowFilterFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdShowFilterStylesAvailable": WdShowFilterFromString = wdShowFilterStylesAvailable
        Case "wdShowFilterStylesInUse": WdShowFilterFromString = wdShowFilterStylesInUse
        Case "wdShowFilterStylesAll": WdShowFilterFromString = wdShowFilterStylesAll
        Case "wdShowFilterFormattingInUse": WdShowFilterFromString = wdShowFilterFormattingInUse
        Case "wdShowFilterFormattingAvailable": WdShowFilterFromString = wdShowFilterFormattingAvailable
        Case "wdShowFilterFormattingRecommended": WdShowFilterFromString = wdShowFilterFormattingRecommended
    End Select
End Function

Function WdShowFilterToString(value As WdShowFilter) As String
    Select Case value
        Case wdShowFilterStylesAvailable: WdShowFilterToString = "wdShowFilterStylesAvailable"
        Case wdShowFilterStylesInUse: WdShowFilterToString = "wdShowFilterStylesInUse"
        Case wdShowFilterStylesAll: WdShowFilterToString = "wdShowFilterStylesAll"
        Case wdShowFilterFormattingInUse: WdShowFilterToString = "wdShowFilterFormattingInUse"
        Case wdShowFilterFormattingAvailable: WdShowFilterToString = "wdShowFilterFormattingAvailable"
        Case wdShowFilterFormattingRecommended: WdShowFilterToString = "wdShowFilterFormattingRecommended"
    End Select
End Function
