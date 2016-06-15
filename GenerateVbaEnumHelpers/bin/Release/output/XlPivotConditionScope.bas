Attribute VB_Name = "wXlPivotConditionScope"
Function XlPivotConditionScopeFromString(value As String) As XlPivotConditionScope
    If IsNumeric(value) Then
        XlPivotConditionScopeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlSelectionScope": XlPivotConditionScopeFromString = xlSelectionScope
        Case "xlFieldsScope": XlPivotConditionScopeFromString = xlFieldsScope
        Case "xlDataFieldScope": XlPivotConditionScopeFromString = xlDataFieldScope
    End Select
End Function

Function XlPivotConditionScopeToString(value As XlPivotConditionScope) As String
    Select Case value
        Case xlSelectionScope: XlPivotConditionScopeToString = "xlSelectionScope"
        Case xlFieldsScope: XlPivotConditionScopeToString = "xlFieldsScope"
        Case xlDataFieldScope: XlPivotConditionScopeToString = "xlDataFieldScope"
    End Select
End Function
