Attribute VB_Name = "wPpUpdateOption"
Function PpUpdateOptionFromString(value As String) As PpUpdateOption
    If IsNumeric(value) Then
        PpUpdateOptionFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "ppUpdateOptionManual": PpUpdateOptionFromString = ppUpdateOptionManual
        Case "ppUpdateOptionAutomatic": PpUpdateOptionFromString = ppUpdateOptionAutomatic
        Case "ppUpdateOptionMixed": PpUpdateOptionFromString = ppUpdateOptionMixed
    End Select
End Function

Function PpUpdateOptionToString(value As PpUpdateOption) As String
    Select Case value
        Case ppUpdateOptionManual: PpUpdateOptionToString = "ppUpdateOptionManual"
        Case ppUpdateOptionAutomatic: PpUpdateOptionToString = "ppUpdateOptionAutomatic"
        Case ppUpdateOptionMixed: PpUpdateOptionToString = "ppUpdateOptionMixed"
    End Select
End Function
