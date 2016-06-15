Attribute VB_Name = "wMsoTextureType"
Function MsoTextureTypeFromString(value As String) As MsoTextureType
    If IsNumeric(value) Then
        MsoTextureTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoTexturePreset": MsoTextureTypeFromString = msoTexturePreset
        Case "msoTextureUserDefined": MsoTextureTypeFromString = msoTextureUserDefined
        Case "msoTextureTypeMixed": MsoTextureTypeFromString = msoTextureTypeMixed
    End Select
End Function

Function MsoTextureTypeToString(value As MsoTextureType) As String
    Select Case value
        Case msoTexturePreset: MsoTextureTypeToString = "msoTexturePreset"
        Case msoTextureUserDefined: MsoTextureTypeToString = "msoTextureUserDefined"
        Case msoTextureTypeMixed: MsoTextureTypeToString = "msoTextureTypeMixed"
    End Select
End Function
