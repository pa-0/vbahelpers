Attribute VB_Name = "wPbLigaturePresetType"
Function PbLigaturePresetTypeFromString(value As String) As PbLigaturePresetType
    If IsNumeric(value) Then
        PbLigaturePresetTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbLigatureStandard": PbLigaturePresetTypeFromString = pbLigatureStandard
        Case "pbLigatureStandardOptional": PbLigaturePresetTypeFromString = pbLigatureStandardOptional
        Case "pbLigatureStandardHistorical": PbLigaturePresetTypeFromString = pbLigatureStandardHistorical
        Case "pbLigatureAll": PbLigaturePresetTypeFromString = pbLigatureAll
        Case "pbLigatureNone": PbLigaturePresetTypeFromString = pbLigatureNone
        Case "pbLigatureMixed": PbLigaturePresetTypeFromString = pbLigatureMixed
    End Select
End Function

Function PbLigaturePresetTypeToString(value As PbLigaturePresetType) As String
    Select Case value
        Case pbLigatureStandard: PbLigaturePresetTypeToString = "pbLigatureStandard"
        Case pbLigatureStandardOptional: PbLigaturePresetTypeToString = "pbLigatureStandardOptional"
        Case pbLigatureStandardHistorical: PbLigaturePresetTypeToString = "pbLigatureStandardHistorical"
        Case pbLigatureAll: PbLigaturePresetTypeToString = "pbLigatureAll"
        Case pbLigatureNone: PbLigaturePresetTypeToString = "pbLigatureNone"
        Case pbLigatureMixed: PbLigaturePresetTypeToString = "pbLigatureMixed"
    End Select
End Function
