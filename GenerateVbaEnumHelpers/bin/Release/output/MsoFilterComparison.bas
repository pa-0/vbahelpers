Attribute VB_Name = "wMsoFilterComparison"
Function MsoFilterComparisonFromString(value As String) As MsoFilterComparison
    If IsNumeric(value) Then
        MsoFilterComparisonFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoFilterComparisonEqual": MsoFilterComparisonFromString = msoFilterComparisonEqual
        Case "msoFilterComparisonNotEqual": MsoFilterComparisonFromString = msoFilterComparisonNotEqual
        Case "msoFilterComparisonLessThan": MsoFilterComparisonFromString = msoFilterComparisonLessThan
        Case "msoFilterComparisonGreaterThan": MsoFilterComparisonFromString = msoFilterComparisonGreaterThan
        Case "msoFilterComparisonLessThanEqual": MsoFilterComparisonFromString = msoFilterComparisonLessThanEqual
        Case "msoFilterComparisonGreaterThanEqual": MsoFilterComparisonFromString = msoFilterComparisonGreaterThanEqual
        Case "msoFilterComparisonIsBlank": MsoFilterComparisonFromString = msoFilterComparisonIsBlank
        Case "msoFilterComparisonIsNotBlank": MsoFilterComparisonFromString = msoFilterComparisonIsNotBlank
        Case "msoFilterComparisonContains": MsoFilterComparisonFromString = msoFilterComparisonContains
        Case "msoFilterComparisonNotContains": MsoFilterComparisonFromString = msoFilterComparisonNotContains
    End Select
End Function

Function MsoFilterComparisonToString(value As MsoFilterComparison) As String
    Select Case value
        Case msoFilterComparisonEqual: MsoFilterComparisonToString = "msoFilterComparisonEqual"
        Case msoFilterComparisonNotEqual: MsoFilterComparisonToString = "msoFilterComparisonNotEqual"
        Case msoFilterComparisonLessThan: MsoFilterComparisonToString = "msoFilterComparisonLessThan"
        Case msoFilterComparisonGreaterThan: MsoFilterComparisonToString = "msoFilterComparisonGreaterThan"
        Case msoFilterComparisonLessThanEqual: MsoFilterComparisonToString = "msoFilterComparisonLessThanEqual"
        Case msoFilterComparisonGreaterThanEqual: MsoFilterComparisonToString = "msoFilterComparisonGreaterThanEqual"
        Case msoFilterComparisonIsBlank: MsoFilterComparisonToString = "msoFilterComparisonIsBlank"
        Case msoFilterComparisonIsNotBlank: MsoFilterComparisonToString = "msoFilterComparisonIsNotBlank"
        Case msoFilterComparisonContains: MsoFilterComparisonToString = "msoFilterComparisonContains"
        Case msoFilterComparisonNotContains: MsoFilterComparisonToString = "msoFilterComparisonNotContains"
    End Select
End Function
