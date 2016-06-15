Attribute VB_Name = "wOlMultiSelect"
Function OlMultiSelectFromString(value As String) As OlMultiSelect
    If IsNumeric(value) Then
        OlMultiSelectFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olMultiSelectSingle": OlMultiSelectFromString = olMultiSelectSingle
        Case "olMultiSelectMulti": OlMultiSelectFromString = olMultiSelectMulti
        Case "olMultiSelectExtended": OlMultiSelectFromString = olMultiSelectExtended
    End Select
End Function

Function OlMultiSelectToString(value As OlMultiSelect) As String
    Select Case value
        Case olMultiSelectSingle: OlMultiSelectToString = "olMultiSelectSingle"
        Case olMultiSelectMulti: OlMultiSelectToString = "olMultiSelectMulti"
        Case olMultiSelectExtended: OlMultiSelectToString = "olMultiSelectExtended"
    End Select
End Function
