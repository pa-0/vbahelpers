Attribute VB_Name = "wWdSummaryMode"
Function WdSummaryModeFromString(value As String) As WdSummaryMode
    If IsNumeric(value) Then
        WdSummaryModeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdSummaryModeHighlight": WdSummaryModeFromString = wdSummaryModeHighlight
        Case "wdSummaryModeHideAllButSummary": WdSummaryModeFromString = wdSummaryModeHideAllButSummary
        Case "wdSummaryModeInsert": WdSummaryModeFromString = wdSummaryModeInsert
        Case "wdSummaryModeCreateNew": WdSummaryModeFromString = wdSummaryModeCreateNew
    End Select
End Function

Function WdSummaryModeToString(value As WdSummaryMode) As String
    Select Case value
        Case wdSummaryModeHighlight: WdSummaryModeToString = "wdSummaryModeHighlight"
        Case wdSummaryModeHideAllButSummary: WdSummaryModeToString = "wdSummaryModeHideAllButSummary"
        Case wdSummaryModeInsert: WdSummaryModeToString = "wdSummaryModeInsert"
        Case wdSummaryModeCreateNew: WdSummaryModeToString = "wdSummaryModeCreateNew"
    End Select
End Function
