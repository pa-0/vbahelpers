Attribute VB_Name = "wMsoPresetGradientType"
Function MsoPresetGradientTypeFromString(value As String) As MsoPresetGradientType
    If IsNumeric(value) Then
        MsoPresetGradientTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoGradientEarlySunset": MsoPresetGradientTypeFromString = msoGradientEarlySunset
        Case "msoGradientLateSunset": MsoPresetGradientTypeFromString = msoGradientLateSunset
        Case "msoGradientNightfall": MsoPresetGradientTypeFromString = msoGradientNightfall
        Case "msoGradientDaybreak": MsoPresetGradientTypeFromString = msoGradientDaybreak
        Case "msoGradientHorizon": MsoPresetGradientTypeFromString = msoGradientHorizon
        Case "msoGradientDesert": MsoPresetGradientTypeFromString = msoGradientDesert
        Case "msoGradientOcean": MsoPresetGradientTypeFromString = msoGradientOcean
        Case "msoGradientCalmWater": MsoPresetGradientTypeFromString = msoGradientCalmWater
        Case "msoGradientFire": MsoPresetGradientTypeFromString = msoGradientFire
        Case "msoGradientFog": MsoPresetGradientTypeFromString = msoGradientFog
        Case "msoGradientMoss": MsoPresetGradientTypeFromString = msoGradientMoss
        Case "msoGradientPeacock": MsoPresetGradientTypeFromString = msoGradientPeacock
        Case "msoGradientWheat": MsoPresetGradientTypeFromString = msoGradientWheat
        Case "msoGradientParchment": MsoPresetGradientTypeFromString = msoGradientParchment
        Case "msoGradientMahogany": MsoPresetGradientTypeFromString = msoGradientMahogany
        Case "msoGradientRainbow": MsoPresetGradientTypeFromString = msoGradientRainbow
        Case "msoGradientRainbowII": MsoPresetGradientTypeFromString = msoGradientRainbowII
        Case "msoGradientGold": MsoPresetGradientTypeFromString = msoGradientGold
        Case "msoGradientGoldII": MsoPresetGradientTypeFromString = msoGradientGoldII
        Case "msoGradientBrass": MsoPresetGradientTypeFromString = msoGradientBrass
        Case "msoGradientChrome": MsoPresetGradientTypeFromString = msoGradientChrome
        Case "msoGradientChromeII": MsoPresetGradientTypeFromString = msoGradientChromeII
        Case "msoGradientSilver": MsoPresetGradientTypeFromString = msoGradientSilver
        Case "msoGradientSapphire": MsoPresetGradientTypeFromString = msoGradientSapphire
        Case "msoPresetGradientMixed": MsoPresetGradientTypeFromString = msoPresetGradientMixed
    End Select
End Function

Function MsoPresetGradientTypeToString(value As MsoPresetGradientType) As String
    Select Case value
        Case msoGradientEarlySunset: MsoPresetGradientTypeToString = "msoGradientEarlySunset"
        Case msoGradientLateSunset: MsoPresetGradientTypeToString = "msoGradientLateSunset"
        Case msoGradientNightfall: MsoPresetGradientTypeToString = "msoGradientNightfall"
        Case msoGradientDaybreak: MsoPresetGradientTypeToString = "msoGradientDaybreak"
        Case msoGradientHorizon: MsoPresetGradientTypeToString = "msoGradientHorizon"
        Case msoGradientDesert: MsoPresetGradientTypeToString = "msoGradientDesert"
        Case msoGradientOcean: MsoPresetGradientTypeToString = "msoGradientOcean"
        Case msoGradientCalmWater: MsoPresetGradientTypeToString = "msoGradientCalmWater"
        Case msoGradientFire: MsoPresetGradientTypeToString = "msoGradientFire"
        Case msoGradientFog: MsoPresetGradientTypeToString = "msoGradientFog"
        Case msoGradientMoss: MsoPresetGradientTypeToString = "msoGradientMoss"
        Case msoGradientPeacock: MsoPresetGradientTypeToString = "msoGradientPeacock"
        Case msoGradientWheat: MsoPresetGradientTypeToString = "msoGradientWheat"
        Case msoGradientParchment: MsoPresetGradientTypeToString = "msoGradientParchment"
        Case msoGradientMahogany: MsoPresetGradientTypeToString = "msoGradientMahogany"
        Case msoGradientRainbow: MsoPresetGradientTypeToString = "msoGradientRainbow"
        Case msoGradientRainbowII: MsoPresetGradientTypeToString = "msoGradientRainbowII"
        Case msoGradientGold: MsoPresetGradientTypeToString = "msoGradientGold"
        Case msoGradientGoldII: MsoPresetGradientTypeToString = "msoGradientGoldII"
        Case msoGradientBrass: MsoPresetGradientTypeToString = "msoGradientBrass"
        Case msoGradientChrome: MsoPresetGradientTypeToString = "msoGradientChrome"
        Case msoGradientChromeII: MsoPresetGradientTypeToString = "msoGradientChromeII"
        Case msoGradientSilver: MsoPresetGradientTypeToString = "msoGradientSilver"
        Case msoGradientSapphire: MsoPresetGradientTypeToString = "msoGradientSapphire"
        Case msoPresetGradientMixed: MsoPresetGradientTypeToString = "msoPresetGradientMixed"
    End Select
End Function
