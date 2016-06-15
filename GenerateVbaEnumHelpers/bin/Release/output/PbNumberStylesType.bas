Attribute VB_Name = "wPbNumberStylesType"
Function PbNumberStylesTypeFromString(value As String) As PbNumberStylesType
    If IsNumeric(value) Then
        PbNumberStylesTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbNumberStyleDefault": PbNumberStylesTypeFromString = pbNumberStyleDefault
        Case "pbNumberStyleProportionalLining": PbNumberStylesTypeFromString = pbNumberStyleProportionalLining
        Case "pbNumberStyleTabularLining": PbNumberStylesTypeFromString = pbNumberStyleTabularLining
        Case "pbNumberStyleProportionalOldstyle": PbNumberStylesTypeFromString = pbNumberStyleProportionalOldstyle
        Case "pbNumberStyleTabularOldstyle": PbNumberStylesTypeFromString = pbNumberStyleTabularOldstyle
        Case "pbNumberStyleMixed": PbNumberStylesTypeFromString = pbNumberStyleMixed
    End Select
End Function

Function PbNumberStylesTypeToString(value As PbNumberStylesType) As String
    Select Case value
        Case pbNumberStyleDefault: PbNumberStylesTypeToString = "pbNumberStyleDefault"
        Case pbNumberStyleProportionalLining: PbNumberStylesTypeToString = "pbNumberStyleProportionalLining"
        Case pbNumberStyleTabularLining: PbNumberStylesTypeToString = "pbNumberStyleTabularLining"
        Case pbNumberStyleProportionalOldstyle: PbNumberStylesTypeToString = "pbNumberStyleProportionalOldstyle"
        Case pbNumberStyleTabularOldstyle: PbNumberStylesTypeToString = "pbNumberStyleTabularOldstyle"
        Case pbNumberStyleMixed: PbNumberStylesTypeToString = "pbNumberStyleMixed"
    End Select
End Function
