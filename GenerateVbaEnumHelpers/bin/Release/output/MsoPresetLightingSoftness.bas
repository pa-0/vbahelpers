Attribute VB_Name = "wMsoPresetLightingSoftness"
Function MsoPresetLightingSoftnessFromString(value As String) As MsoPresetLightingSoftness
    If IsNumeric(value) Then
        MsoPresetLightingSoftnessFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoLightingDim": MsoPresetLightingSoftnessFromString = msoLightingDim
        Case "msoLightingNormal": MsoPresetLightingSoftnessFromString = msoLightingNormal
        Case "msoLightingBright": MsoPresetLightingSoftnessFromString = msoLightingBright
        Case "msoPresetLightingSoftnessMixed": MsoPresetLightingSoftnessFromString = msoPresetLightingSoftnessMixed
    End Select
End Function

Function MsoPresetLightingSoftnessToString(value As MsoPresetLightingSoftness) As String
    Select Case value
        Case msoLightingDim: MsoPresetLightingSoftnessToString = "msoLightingDim"
        Case msoLightingNormal: MsoPresetLightingSoftnessToString = "msoLightingNormal"
        Case msoLightingBright: MsoPresetLightingSoftnessToString = "msoLightingBright"
        Case msoPresetLightingSoftnessMixed: MsoPresetLightingSoftnessToString = "msoPresetLightingSoftnessMixed"
    End Select
End Function
