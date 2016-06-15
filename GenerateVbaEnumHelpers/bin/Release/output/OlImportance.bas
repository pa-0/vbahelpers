Attribute VB_Name = "wOlImportance"
Function OlImportanceFromString(value As String) As OlImportance
    If IsNumeric(value) Then
        OlImportanceFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olImportanceLow": OlImportanceFromString = olImportanceLow
        Case "olImportanceNormal": OlImportanceFromString = olImportanceNormal
        Case "olImportanceHigh": OlImportanceFromString = olImportanceHigh
    End Select
End Function

Function OlImportanceToString(value As OlImportance) As String
    Select Case value
        Case olImportanceLow: OlImportanceToString = "olImportanceLow"
        Case olImportanceNormal: OlImportanceToString = "olImportanceNormal"
        Case olImportanceHigh: OlImportanceToString = "olImportanceHigh"
    End Select
End Function
