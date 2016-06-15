Attribute VB_Name = "wMsoPictureEffectType"
Function MsoPictureEffectTypeFromString(value As String) As MsoPictureEffectType
    If IsNumeric(value) Then
        MsoPictureEffectTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoEffectNone": MsoPictureEffectTypeFromString = msoEffectNone
        Case "msoEffectBackgroundRemoval": MsoPictureEffectTypeFromString = msoEffectBackgroundRemoval
        Case "msoEffectBlur": MsoPictureEffectTypeFromString = msoEffectBlur
        Case "msoEffectBrightnessContrast": MsoPictureEffectTypeFromString = msoEffectBrightnessContrast
        Case "msoEffectCement": MsoPictureEffectTypeFromString = msoEffectCement
        Case "msoEffectCrisscrossEtching": MsoPictureEffectTypeFromString = msoEffectCrisscrossEtching
        Case "msoEffectChalkSketch": MsoPictureEffectTypeFromString = msoEffectChalkSketch
        Case "msoEffectColorTemperature": MsoPictureEffectTypeFromString = msoEffectColorTemperature
        Case "msoEffectCutout": MsoPictureEffectTypeFromString = msoEffectCutout
        Case "msoEffectFilmGrain": MsoPictureEffectTypeFromString = msoEffectFilmGrain
        Case "msoEffectGlass": MsoPictureEffectTypeFromString = msoEffectGlass
        Case "msoEffectGlowDiffused": MsoPictureEffectTypeFromString = msoEffectGlowDiffused
        Case "msoEffectGlowEdges": MsoPictureEffectTypeFromString = msoEffectGlowEdges
        Case "msoEffectLightScreen": MsoPictureEffectTypeFromString = msoEffectLightScreen
        Case "msoEffectLineDrawing": MsoPictureEffectTypeFromString = msoEffectLineDrawing
        Case "msoEffectMarker": MsoPictureEffectTypeFromString = msoEffectMarker
        Case "msoEffectMosiaicBubbles": MsoPictureEffectTypeFromString = msoEffectMosiaicBubbles
        Case "msoEffectPaintBrush": MsoPictureEffectTypeFromString = msoEffectPaintBrush
        Case "msoEffectPaintStrokes": MsoPictureEffectTypeFromString = msoEffectPaintStrokes
        Case "msoEffectPastelsSmooth": MsoPictureEffectTypeFromString = msoEffectPastelsSmooth
        Case "msoEffectPencilGrayscale": MsoPictureEffectTypeFromString = msoEffectPencilGrayscale
        Case "msoEffectPencilSketch": MsoPictureEffectTypeFromString = msoEffectPencilSketch
        Case "msoEffectPhotocopy": MsoPictureEffectTypeFromString = msoEffectPhotocopy
        Case "msoEffectPlasticWrap": MsoPictureEffectTypeFromString = msoEffectPlasticWrap
        Case "msoEffectSaturation": MsoPictureEffectTypeFromString = msoEffectSaturation
        Case "msoEffectSharpenSoften": MsoPictureEffectTypeFromString = msoEffectSharpenSoften
        Case "msoEffectTexturizer": MsoPictureEffectTypeFromString = msoEffectTexturizer
        Case "msoEffectWatercolorSponge": MsoPictureEffectTypeFromString = msoEffectWatercolorSponge
    End Select
End Function

Function MsoPictureEffectTypeToString(value As MsoPictureEffectType) As String
    Select Case value
        Case msoEffectNone: MsoPictureEffectTypeToString = "msoEffectNone"
        Case msoEffectBackgroundRemoval: MsoPictureEffectTypeToString = "msoEffectBackgroundRemoval"
        Case msoEffectBlur: MsoPictureEffectTypeToString = "msoEffectBlur"
        Case msoEffectBrightnessContrast: MsoPictureEffectTypeToString = "msoEffectBrightnessContrast"
        Case msoEffectCement: MsoPictureEffectTypeToString = "msoEffectCement"
        Case msoEffectCrisscrossEtching: MsoPictureEffectTypeToString = "msoEffectCrisscrossEtching"
        Case msoEffectChalkSketch: MsoPictureEffectTypeToString = "msoEffectChalkSketch"
        Case msoEffectColorTemperature: MsoPictureEffectTypeToString = "msoEffectColorTemperature"
        Case msoEffectCutout: MsoPictureEffectTypeToString = "msoEffectCutout"
        Case msoEffectFilmGrain: MsoPictureEffectTypeToString = "msoEffectFilmGrain"
        Case msoEffectGlass: MsoPictureEffectTypeToString = "msoEffectGlass"
        Case msoEffectGlowDiffused: MsoPictureEffectTypeToString = "msoEffectGlowDiffused"
        Case msoEffectGlowEdges: MsoPictureEffectTypeToString = "msoEffectGlowEdges"
        Case msoEffectLightScreen: MsoPictureEffectTypeToString = "msoEffectLightScreen"
        Case msoEffectLineDrawing: MsoPictureEffectTypeToString = "msoEffectLineDrawing"
        Case msoEffectMarker: MsoPictureEffectTypeToString = "msoEffectMarker"
        Case msoEffectMosiaicBubbles: MsoPictureEffectTypeToString = "msoEffectMosiaicBubbles"
        Case msoEffectPaintBrush: MsoPictureEffectTypeToString = "msoEffectPaintBrush"
        Case msoEffectPaintStrokes: MsoPictureEffectTypeToString = "msoEffectPaintStrokes"
        Case msoEffectPastelsSmooth: MsoPictureEffectTypeToString = "msoEffectPastelsSmooth"
        Case msoEffectPencilGrayscale: MsoPictureEffectTypeToString = "msoEffectPencilGrayscale"
        Case msoEffectPencilSketch: MsoPictureEffectTypeToString = "msoEffectPencilSketch"
        Case msoEffectPhotocopy: MsoPictureEffectTypeToString = "msoEffectPhotocopy"
        Case msoEffectPlasticWrap: MsoPictureEffectTypeToString = "msoEffectPlasticWrap"
        Case msoEffectSaturation: MsoPictureEffectTypeToString = "msoEffectSaturation"
        Case msoEffectSharpenSoften: MsoPictureEffectTypeToString = "msoEffectSharpenSoften"
        Case msoEffectTexturizer: MsoPictureEffectTypeToString = "msoEffectTexturizer"
        Case msoEffectWatercolorSponge: MsoPictureEffectTypeToString = "msoEffectWatercolorSponge"
    End Select
End Function
