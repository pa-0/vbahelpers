Attribute VB_Name = "wPbTrackingPresetType"
Function PbTrackingPresetTypeFromString(value As String) As PbTrackingPresetType
    If IsNumeric(value) Then
        PbTrackingPresetTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbTrackingVeryLoose": PbTrackingPresetTypeFromString = pbTrackingVeryLoose
        Case "pbTrackingLoose": PbTrackingPresetTypeFromString = pbTrackingLoose
        Case "pbTrackingNormal": PbTrackingPresetTypeFromString = pbTrackingNormal
        Case "pbTrackingTight": PbTrackingPresetTypeFromString = pbTrackingTight
        Case "pbTrackingVeryTight": PbTrackingPresetTypeFromString = pbTrackingVeryTight
        Case "pbTrackingMixed": PbTrackingPresetTypeFromString = pbTrackingMixed
        Case "pbTrackingCustom": PbTrackingPresetTypeFromString = pbTrackingCustom
    End Select
End Function

Function PbTrackingPresetTypeToString(value As PbTrackingPresetType) As String
    Select Case value
        Case pbTrackingVeryLoose: PbTrackingPresetTypeToString = "pbTrackingVeryLoose"
        Case pbTrackingLoose: PbTrackingPresetTypeToString = "pbTrackingLoose"
        Case pbTrackingNormal: PbTrackingPresetTypeToString = "pbTrackingNormal"
        Case pbTrackingTight: PbTrackingPresetTypeToString = "pbTrackingTight"
        Case pbTrackingVeryTight: PbTrackingPresetTypeToString = "pbTrackingVeryTight"
        Case pbTrackingMixed: PbTrackingPresetTypeToString = "pbTrackingMixed"
        Case pbTrackingCustom: PbTrackingPresetTypeToString = "pbTrackingCustom"
    End Select
End Function
