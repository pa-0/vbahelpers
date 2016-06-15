Attribute VB_Name = "wMsoAnimFilterEffectType"
Function MsoAnimFilterEffectTypeFromString(value As String) As MsoAnimFilterEffectType
    If IsNumeric(value) Then
        MsoAnimFilterEffectTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoAnimFilterEffectTypeNone": MsoAnimFilterEffectTypeFromString = msoAnimFilterEffectTypeNone
        Case "msoAnimFilterEffectTypeBarn": MsoAnimFilterEffectTypeFromString = msoAnimFilterEffectTypeBarn
        Case "msoAnimFilterEffectTypeBlinds": MsoAnimFilterEffectTypeFromString = msoAnimFilterEffectTypeBlinds
        Case "msoAnimFilterEffectTypeBox": MsoAnimFilterEffectTypeFromString = msoAnimFilterEffectTypeBox
        Case "msoAnimFilterEffectTypeCheckerboard": MsoAnimFilterEffectTypeFromString = msoAnimFilterEffectTypeCheckerboard
        Case "msoAnimFilterEffectTypeCircle": MsoAnimFilterEffectTypeFromString = msoAnimFilterEffectTypeCircle
        Case "msoAnimFilterEffectTypeDiamond": MsoAnimFilterEffectTypeFromString = msoAnimFilterEffectTypeDiamond
        Case "msoAnimFilterEffectTypeDissolve": MsoAnimFilterEffectTypeFromString = msoAnimFilterEffectTypeDissolve
        Case "msoAnimFilterEffectTypeFade": MsoAnimFilterEffectTypeFromString = msoAnimFilterEffectTypeFade
        Case "msoAnimFilterEffectTypeImage": MsoAnimFilterEffectTypeFromString = msoAnimFilterEffectTypeImage
        Case "msoAnimFilterEffectTypePixelate": MsoAnimFilterEffectTypeFromString = msoAnimFilterEffectTypePixelate
        Case "msoAnimFilterEffectTypePlus": MsoAnimFilterEffectTypeFromString = msoAnimFilterEffectTypePlus
        Case "msoAnimFilterEffectTypeRandomBar": MsoAnimFilterEffectTypeFromString = msoAnimFilterEffectTypeRandomBar
        Case "msoAnimFilterEffectTypeSlide": MsoAnimFilterEffectTypeFromString = msoAnimFilterEffectTypeSlide
        Case "msoAnimFilterEffectTypeStretch": MsoAnimFilterEffectTypeFromString = msoAnimFilterEffectTypeStretch
        Case "msoAnimFilterEffectTypeStrips": MsoAnimFilterEffectTypeFromString = msoAnimFilterEffectTypeStrips
        Case "msoAnimFilterEffectTypeWedge": MsoAnimFilterEffectTypeFromString = msoAnimFilterEffectTypeWedge
        Case "msoAnimFilterEffectTypeWheel": MsoAnimFilterEffectTypeFromString = msoAnimFilterEffectTypeWheel
        Case "msoAnimFilterEffectTypeWipe": MsoAnimFilterEffectTypeFromString = msoAnimFilterEffectTypeWipe
    End Select
End Function

Function MsoAnimFilterEffectTypeToString(value As MsoAnimFilterEffectType) As String
    Select Case value
        Case msoAnimFilterEffectTypeNone: MsoAnimFilterEffectTypeToString = "msoAnimFilterEffectTypeNone"
        Case msoAnimFilterEffectTypeBarn: MsoAnimFilterEffectTypeToString = "msoAnimFilterEffectTypeBarn"
        Case msoAnimFilterEffectTypeBlinds: MsoAnimFilterEffectTypeToString = "msoAnimFilterEffectTypeBlinds"
        Case msoAnimFilterEffectTypeBox: MsoAnimFilterEffectTypeToString = "msoAnimFilterEffectTypeBox"
        Case msoAnimFilterEffectTypeCheckerboard: MsoAnimFilterEffectTypeToString = "msoAnimFilterEffectTypeCheckerboard"
        Case msoAnimFilterEffectTypeCircle: MsoAnimFilterEffectTypeToString = "msoAnimFilterEffectTypeCircle"
        Case msoAnimFilterEffectTypeDiamond: MsoAnimFilterEffectTypeToString = "msoAnimFilterEffectTypeDiamond"
        Case msoAnimFilterEffectTypeDissolve: MsoAnimFilterEffectTypeToString = "msoAnimFilterEffectTypeDissolve"
        Case msoAnimFilterEffectTypeFade: MsoAnimFilterEffectTypeToString = "msoAnimFilterEffectTypeFade"
        Case msoAnimFilterEffectTypeImage: MsoAnimFilterEffectTypeToString = "msoAnimFilterEffectTypeImage"
        Case msoAnimFilterEffectTypePixelate: MsoAnimFilterEffectTypeToString = "msoAnimFilterEffectTypePixelate"
        Case msoAnimFilterEffectTypePlus: MsoAnimFilterEffectTypeToString = "msoAnimFilterEffectTypePlus"
        Case msoAnimFilterEffectTypeRandomBar: MsoAnimFilterEffectTypeToString = "msoAnimFilterEffectTypeRandomBar"
        Case msoAnimFilterEffectTypeSlide: MsoAnimFilterEffectTypeToString = "msoAnimFilterEffectTypeSlide"
        Case msoAnimFilterEffectTypeStretch: MsoAnimFilterEffectTypeToString = "msoAnimFilterEffectTypeStretch"
        Case msoAnimFilterEffectTypeStrips: MsoAnimFilterEffectTypeToString = "msoAnimFilterEffectTypeStrips"
        Case msoAnimFilterEffectTypeWedge: MsoAnimFilterEffectTypeToString = "msoAnimFilterEffectTypeWedge"
        Case msoAnimFilterEffectTypeWheel: MsoAnimFilterEffectTypeToString = "msoAnimFilterEffectTypeWheel"
        Case msoAnimFilterEffectTypeWipe: MsoAnimFilterEffectTypeToString = "msoAnimFilterEffectTypeWipe"
    End Select
End Function
