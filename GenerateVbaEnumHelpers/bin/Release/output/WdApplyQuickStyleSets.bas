Attribute VB_Name = "wWdApplyQuickStyleSets"
Function WdApplyQuickStyleSetsFromString(value As String) As WdApplyQuickStyleSets
    If IsNumeric(value) Then
        WdApplyQuickStyleSetsFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdSessionStartSet": WdApplyQuickStyleSetsFromString = wdSessionStartSet
        Case "wdTemplateSet": WdApplyQuickStyleSetsFromString = wdTemplateSet
    End Select
End Function

Function WdApplyQuickStyleSetsToString(value As WdApplyQuickStyleSets) As String
    Select Case value
        Case wdSessionStartSet: WdApplyQuickStyleSetsToString = "wdSessionStartSet"
        Case wdTemplateSet: WdApplyQuickStyleSetsToString = "wdTemplateSet"
    End Select
End Function
