Attribute VB_Name = "wMsoPresetTexture"
Function MsoPresetTextureFromString(value As String) As MsoPresetTexture
    If IsNumeric(value) Then
        MsoPresetTextureFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoTexturePapyrus": MsoPresetTextureFromString = msoTexturePapyrus
        Case "msoTextureCanvas": MsoPresetTextureFromString = msoTextureCanvas
        Case "msoTextureDenim": MsoPresetTextureFromString = msoTextureDenim
        Case "msoTextureWovenMat": MsoPresetTextureFromString = msoTextureWovenMat
        Case "msoTextureWaterDroplets": MsoPresetTextureFromString = msoTextureWaterDroplets
        Case "msoTexturePaperBag": MsoPresetTextureFromString = msoTexturePaperBag
        Case "msoTextureFishFossil": MsoPresetTextureFromString = msoTextureFishFossil
        Case "msoTextureSand": MsoPresetTextureFromString = msoTextureSand
        Case "msoTextureGreenMarble": MsoPresetTextureFromString = msoTextureGreenMarble
        Case "msoTextureWhiteMarble": MsoPresetTextureFromString = msoTextureWhiteMarble
        Case "msoTextureBrownMarble": MsoPresetTextureFromString = msoTextureBrownMarble
        Case "msoTextureGranite": MsoPresetTextureFromString = msoTextureGranite
        Case "msoTextureNewsprint": MsoPresetTextureFromString = msoTextureNewsprint
        Case "msoTextureRecycledPaper": MsoPresetTextureFromString = msoTextureRecycledPaper
        Case "msoTextureParchment": MsoPresetTextureFromString = msoTextureParchment
        Case "msoTextureStationery": MsoPresetTextureFromString = msoTextureStationery
        Case "msoTextureBlueTissuePaper": MsoPresetTextureFromString = msoTextureBlueTissuePaper
        Case "msoTexturePinkTissuePaper": MsoPresetTextureFromString = msoTexturePinkTissuePaper
        Case "msoTexturePurpleMesh": MsoPresetTextureFromString = msoTexturePurpleMesh
        Case "msoTextureBouquet": MsoPresetTextureFromString = msoTextureBouquet
        Case "msoTextureCork": MsoPresetTextureFromString = msoTextureCork
        Case "msoTextureWalnut": MsoPresetTextureFromString = msoTextureWalnut
        Case "msoTextureOak": MsoPresetTextureFromString = msoTextureOak
        Case "msoTextureMediumWood": MsoPresetTextureFromString = msoTextureMediumWood
        Case "msoPresetTextureMixed": MsoPresetTextureFromString = msoPresetTextureMixed
    End Select
End Function

Function MsoPresetTextureToString(value As MsoPresetTexture) As String
    Select Case value
        Case msoTexturePapyrus: MsoPresetTextureToString = "msoTexturePapyrus"
        Case msoTextureCanvas: MsoPresetTextureToString = "msoTextureCanvas"
        Case msoTextureDenim: MsoPresetTextureToString = "msoTextureDenim"
        Case msoTextureWovenMat: MsoPresetTextureToString = "msoTextureWovenMat"
        Case msoTextureWaterDroplets: MsoPresetTextureToString = "msoTextureWaterDroplets"
        Case msoTexturePaperBag: MsoPresetTextureToString = "msoTexturePaperBag"
        Case msoTextureFishFossil: MsoPresetTextureToString = "msoTextureFishFossil"
        Case msoTextureSand: MsoPresetTextureToString = "msoTextureSand"
        Case msoTextureGreenMarble: MsoPresetTextureToString = "msoTextureGreenMarble"
        Case msoTextureWhiteMarble: MsoPresetTextureToString = "msoTextureWhiteMarble"
        Case msoTextureBrownMarble: MsoPresetTextureToString = "msoTextureBrownMarble"
        Case msoTextureGranite: MsoPresetTextureToString = "msoTextureGranite"
        Case msoTextureNewsprint: MsoPresetTextureToString = "msoTextureNewsprint"
        Case msoTextureRecycledPaper: MsoPresetTextureToString = "msoTextureRecycledPaper"
        Case msoTextureParchment: MsoPresetTextureToString = "msoTextureParchment"
        Case msoTextureStationery: MsoPresetTextureToString = "msoTextureStationery"
        Case msoTextureBlueTissuePaper: MsoPresetTextureToString = "msoTextureBlueTissuePaper"
        Case msoTexturePinkTissuePaper: MsoPresetTextureToString = "msoTexturePinkTissuePaper"
        Case msoTexturePurpleMesh: MsoPresetTextureToString = "msoTexturePurpleMesh"
        Case msoTextureBouquet: MsoPresetTextureToString = "msoTextureBouquet"
        Case msoTextureCork: MsoPresetTextureToString = "msoTextureCork"
        Case msoTextureWalnut: MsoPresetTextureToString = "msoTextureWalnut"
        Case msoTextureOak: MsoPresetTextureToString = "msoTextureOak"
        Case msoTextureMediumWood: MsoPresetTextureToString = "msoTextureMediumWood"
        Case msoPresetTextureMixed: MsoPresetTextureToString = "msoPresetTextureMixed"
    End Select
End Function
