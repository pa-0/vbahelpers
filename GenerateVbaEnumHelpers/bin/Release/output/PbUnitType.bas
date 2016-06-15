Attribute VB_Name = "wPbUnitType"
Function PbUnitTypeFromString(value As String) As PbUnitType
    If IsNumeric(value) Then
        PbUnitTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbUnitInch": PbUnitTypeFromString = pbUnitInch
        Case "pbUnitCM": PbUnitTypeFromString = pbUnitCM
        Case "pbUnitPica": PbUnitTypeFromString = pbUnitPica
        Case "pbUnitPoint": PbUnitTypeFromString = pbUnitPoint
        Case "pbUnitEmu": PbUnitTypeFromString = pbUnitEmu
        Case "pbUnitTwip": PbUnitTypeFromString = pbUnitTwip
        Case "pbUnitFeet": PbUnitTypeFromString = pbUnitFeet
        Case "pbUnitMeter": PbUnitTypeFromString = pbUnitMeter
        Case "pbUnitKyu": PbUnitTypeFromString = pbUnitKyu
        Case "pbUnitHa": PbUnitTypeFromString = pbUnitHa
        Case "pbUnitPixel": PbUnitTypeFromString = pbUnitPixel
    End Select
End Function

Function PbUnitTypeToString(value As PbUnitType) As String
    Select Case value
        Case pbUnitInch: PbUnitTypeToString = "pbUnitInch"
        Case pbUnitCM: PbUnitTypeToString = "pbUnitCM"
        Case pbUnitPica: PbUnitTypeToString = "pbUnitPica"
        Case pbUnitPoint: PbUnitTypeToString = "pbUnitPoint"
        Case pbUnitEmu: PbUnitTypeToString = "pbUnitEmu"
        Case pbUnitTwip: PbUnitTypeToString = "pbUnitTwip"
        Case pbUnitFeet: PbUnitTypeToString = "pbUnitFeet"
        Case pbUnitMeter: PbUnitTypeToString = "pbUnitMeter"
        Case pbUnitKyu: PbUnitTypeToString = "pbUnitKyu"
        Case pbUnitHa: PbUnitTypeToString = "pbUnitHa"
        Case pbUnitPixel: PbUnitTypeToString = "pbUnitPixel"
    End Select
End Function
