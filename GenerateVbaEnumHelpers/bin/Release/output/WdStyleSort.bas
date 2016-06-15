Attribute VB_Name = "wWdStyleSort"
Function WdStyleSortFromString(value As String) As WdStyleSort
    If IsNumeric(value) Then
        WdStyleSortFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdStyleSortByName": WdStyleSortFromString = wdStyleSortByName
        Case "wdStyleSortRecommended": WdStyleSortFromString = wdStyleSortRecommended
        Case "wdStyleSortByFont": WdStyleSortFromString = wdStyleSortByFont
        Case "wdStyleSortByBasedOn": WdStyleSortFromString = wdStyleSortByBasedOn
        Case "wdStyleSortByType": WdStyleSortFromString = wdStyleSortByType
    End Select
End Function

Function WdStyleSortToString(value As WdStyleSort) As String
    Select Case value
        Case wdStyleSortByName: WdStyleSortToString = "wdStyleSortByName"
        Case wdStyleSortRecommended: WdStyleSortToString = "wdStyleSortRecommended"
        Case wdStyleSortByFont: WdStyleSortToString = "wdStyleSortByFont"
        Case wdStyleSortByBasedOn: WdStyleSortToString = "wdStyleSortByBasedOn"
        Case wdStyleSortByType: WdStyleSortToString = "wdStyleSortByType"
    End Select
End Function
