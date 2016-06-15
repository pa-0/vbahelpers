Attribute VB_Name = "wOlComboBoxStyle"
Function OlComboBoxStyleFromString(value As String) As OlComboBoxStyle
    If IsNumeric(value) Then
        OlComboBoxStyleFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olComboBoxStyleCombo": OlComboBoxStyleFromString = olComboBoxStyleCombo
        Case "olComboBoxStyleList": OlComboBoxStyleFromString = olComboBoxStyleList
    End Select
End Function

Function OlComboBoxStyleToString(value As OlComboBoxStyle) As String
    Select Case value
        Case olComboBoxStyleCombo: OlComboBoxStyleToString = "olComboBoxStyleCombo"
        Case olComboBoxStyleList: OlComboBoxStyleToString = "olComboBoxStyleList"
    End Select
End Function
