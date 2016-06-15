Attribute VB_Name = "wWdMailMergeComparison"
Function WdMailMergeComparisonFromString(value As String) As WdMailMergeComparison
    If IsNumeric(value) Then
        WdMailMergeComparisonFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdMergeIfEqual": WdMailMergeComparisonFromString = wdMergeIfEqual
        Case "wdMergeIfNotEqual": WdMailMergeComparisonFromString = wdMergeIfNotEqual
        Case "wdMergeIfLessThan": WdMailMergeComparisonFromString = wdMergeIfLessThan
        Case "wdMergeIfGreaterThan": WdMailMergeComparisonFromString = wdMergeIfGreaterThan
        Case "wdMergeIfLessThanOrEqual": WdMailMergeComparisonFromString = wdMergeIfLessThanOrEqual
        Case "wdMergeIfGreaterThanOrEqual": WdMailMergeComparisonFromString = wdMergeIfGreaterThanOrEqual
        Case "wdMergeIfIsBlank": WdMailMergeComparisonFromString = wdMergeIfIsBlank
        Case "wdMergeIfIsNotBlank": WdMailMergeComparisonFromString = wdMergeIfIsNotBlank
    End Select
End Function

Function WdMailMergeComparisonToString(value As WdMailMergeComparison) As String
    Select Case value
        Case wdMergeIfEqual: WdMailMergeComparisonToString = "wdMergeIfEqual"
        Case wdMergeIfNotEqual: WdMailMergeComparisonToString = "wdMergeIfNotEqual"
        Case wdMergeIfLessThan: WdMailMergeComparisonToString = "wdMergeIfLessThan"
        Case wdMergeIfGreaterThan: WdMailMergeComparisonToString = "wdMergeIfGreaterThan"
        Case wdMergeIfLessThanOrEqual: WdMailMergeComparisonToString = "wdMergeIfLessThanOrEqual"
        Case wdMergeIfGreaterThanOrEqual: WdMailMergeComparisonToString = "wdMergeIfGreaterThanOrEqual"
        Case wdMergeIfIsBlank: WdMailMergeComparisonToString = "wdMergeIfIsBlank"
        Case wdMergeIfIsNotBlank: WdMailMergeComparisonToString = "wdMergeIfIsNotBlank"
    End Select
End Function
