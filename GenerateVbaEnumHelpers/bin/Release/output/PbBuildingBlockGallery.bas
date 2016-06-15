Attribute VB_Name = "wPbBuildingBlockGallery"
Function PbBuildingBlockGalleryFromString(value As String) As PbBuildingBlockGallery
    If IsNumeric(value) Then
        PbBuildingBlockGalleryFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbBBGalAdvertisements": PbBuildingBlockGalleryFromString = pbBBGalAdvertisements
        Case "pbBBGalAccents": PbBuildingBlockGalleryFromString = pbBBGalAccents
        Case "pbBBGalCalendars": PbBuildingBlockGalleryFromString = pbBBGalCalendars
        Case "pbBBGalBusinessInfo": PbBuildingBlockGalleryFromString = pbBBGalBusinessInfo
        Case "pbBBGalPageParts": PbBuildingBlockGalleryFromString = pbBBGalPageParts
        Case "pbBBGalNone": PbBuildingBlockGalleryFromString = pbBBGalNone
    End Select
End Function

Function PbBuildingBlockGalleryToString(value As PbBuildingBlockGallery) As String
    Select Case value
        Case pbBBGalAdvertisements: PbBuildingBlockGalleryToString = "pbBBGalAdvertisements"
        Case pbBBGalAccents: PbBuildingBlockGalleryToString = "pbBBGalAccents"
        Case pbBBGalCalendars: PbBuildingBlockGalleryToString = "pbBBGalCalendars"
        Case pbBBGalBusinessInfo: PbBuildingBlockGalleryToString = "pbBBGalBusinessInfo"
        Case pbBBGalPageParts: PbBuildingBlockGalleryToString = "pbBBGalPageParts"
        Case pbBBGalNone: PbBuildingBlockGalleryToString = "pbBBGalNone"
    End Select
End Function
