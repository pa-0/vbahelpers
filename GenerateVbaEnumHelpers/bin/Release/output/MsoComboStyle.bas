Attribute VB_Name = "wMsoComboStyle"
Function MsoComboStyleFromString(value As String) As MsoComboStyle
    If IsNumeric(value) Then
        MsoComboStyleFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoComboNormal": MsoComboStyleFromString = msoComboNormal
        Case "msoComboLabel": MsoComboStyleFromString = msoComboLabel
    End Select
End Function

Function MsoComboStyleToString(value As MsoComboStyle) As String
    Select Case value
        Case msoComboNormal: MsoComboStyleToString = "msoComboNormal"
        Case msoComboLabel: MsoComboStyleToString = "msoComboLabel"
    End Select
End Function
