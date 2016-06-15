Attribute VB_Name = "wPbFilterComparison"
Function PbFilterComparisonFromString(value As String) As PbFilterComparison
    If IsNumeric(value) Then
        PbFilterComparisonFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbComparisonEqual": PbFilterComparisonFromString = pbComparisonEqual
        Case "pbComparisonNotEqual": PbFilterComparisonFromString = pbComparisonNotEqual
        Case "pbComparisonLessThan": PbFilterComparisonFromString = pbComparisonLessThan
        Case "pbComparisonGreaterThan": PbFilterComparisonFromString = pbComparisonGreaterThan
        Case "pbComparisonLessThanEqual": PbFilterComparisonFromString = pbComparisonLessThanEqual
        Case "pbComparisonGreaterThanEqual": PbFilterComparisonFromString = pbComparisonGreaterThanEqual
        Case "pbComparisonIsBlank": PbFilterComparisonFromString = pbComparisonIsBlank
        Case "pbComparisonIsNotBlank": PbFilterComparisonFromString = pbComparisonIsNotBlank
    End Select
End Function

Function PbFilterComparisonToString(value As PbFilterComparison) As String
    Select Case value
        Case pbComparisonEqual: PbFilterComparisonToString = "pbComparisonEqual"
        Case pbComparisonNotEqual: PbFilterComparisonToString = "pbComparisonNotEqual"
        Case pbComparisonLessThan: PbFilterComparisonToString = "pbComparisonLessThan"
        Case pbComparisonGreaterThan: PbFilterComparisonToString = "pbComparisonGreaterThan"
        Case pbComparisonLessThanEqual: PbFilterComparisonToString = "pbComparisonLessThanEqual"
        Case pbComparisonGreaterThanEqual: PbFilterComparisonToString = "pbComparisonGreaterThanEqual"
        Case pbComparisonIsBlank: PbFilterComparisonToString = "pbComparisonIsBlank"
        Case pbComparisonIsNotBlank: PbFilterComparisonToString = "pbComparisonIsNotBlank"
    End Select
End Function
