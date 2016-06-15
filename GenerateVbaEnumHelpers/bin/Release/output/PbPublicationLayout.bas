Attribute VB_Name = "wPbPublicationLayout"
Function PbPublicationLayoutFromString(value As String) As PbPublicationLayout
    If IsNumeric(value) Then
        PbPublicationLayoutFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbLayoutFullPage": PbPublicationLayoutFromString = pbLayoutFullPage
        Case "pbLayoutBook": PbPublicationLayoutFromString = pbLayoutBook
        Case "pbLayoutFoldCard": PbPublicationLayoutFromString = pbLayoutFoldCard
        Case "pbLayoutGreetingCardL": PbPublicationLayoutFromString = pbLayoutGreetingCardL
        Case "pbLayoutGreetingCardT": PbPublicationLayoutFromString = pbLayoutGreetingCardT
        Case "pbLayoutKookBaePan": PbPublicationLayoutFromString = pbLayoutKookBaePan
        Case "pbLayoutKookPan": PbPublicationLayoutFromString = pbLayoutKookPan
        Case "pbLayoutShinKookPan": PbPublicationLayoutFromString = pbLayoutShinKookPan
        Case "pbLayoutKookBanPan": PbPublicationLayoutFromString = pbLayoutKookBanPan
        Case "pbLayout4x6BaePan": PbPublicationLayoutFromString = pbLayout4x6BaePan
        Case "pbLayout4x6Pan": PbPublicationLayoutFromString = pbLayout4x6Pan
        Case "pbLayout4x6BanPan": PbPublicationLayoutFromString = pbLayout4x6BanPan
        Case "pbLayoutCrownPan": PbPublicationLayoutFromString = pbLayoutCrownPan
        Case "pbLayoutShinSeoPan": PbPublicationLayoutFromString = pbLayoutShinSeoPan
        Case "pbLayoutJang4x6Pan": PbPublicationLayoutFromString = pbLayoutJang4x6Pan
        Case "pbLayoutIndexCard": PbPublicationLayoutFromString = pbLayoutIndexCard
        Case "pbLayoutBusinessCardUS": PbPublicationLayoutFromString = pbLayoutBusinessCardUS
        Case "pbLayoutBusinessCardEurope": PbPublicationLayoutFromString = pbLayoutBusinessCardEurope
        Case "pbLayoutBusinessCardFE": PbPublicationLayoutFromString = pbLayoutBusinessCardFE
        Case "pbLayoutBusinessCardLocal": PbPublicationLayoutFromString = pbLayoutBusinessCardLocal
        Case "pbLayoutPosterSmall": PbPublicationLayoutFromString = pbLayoutPosterSmall
        Case "pbLayoutPosterLarge": PbPublicationLayoutFromString = pbLayoutPosterLarge
        Case "pbLayoutCustom": PbPublicationLayoutFromString = pbLayoutCustom
        Case "pbLayoutBannerSmall": PbPublicationLayoutFromString = pbLayoutBannerSmall
        Case "pbLayoutBannerMedium": PbPublicationLayoutFromString = pbLayoutBannerMedium
        Case "pbLayoutBannerLarge": PbPublicationLayoutFromString = pbLayoutBannerLarge
        Case "pbLayoutBannerCustom": PbPublicationLayoutFromString = pbLayoutBannerCustom
        Case "pbLayoutPostcardUS": PbPublicationLayoutFromString = pbLayoutPostcardUS
        Case "pbLayoutPostcardHalfLetter": PbPublicationLayoutFromString = pbLayoutPostcardHalfLetter
        Case "pbLayoutPostcardA4": PbPublicationLayoutFromString = pbLayoutPostcardA4
        Case "pbLayoutPostcardJapan": PbPublicationLayoutFromString = pbLayoutPostcardJapan
        Case "pbLayoutLabel": PbPublicationLayoutFromString = pbLayoutLabel
        Case "pbLayoutEnvelope": PbPublicationLayoutFromString = pbLayoutEnvelope
        Case "pbLayoutWebPageSmall": PbPublicationLayoutFromString = pbLayoutWebPageSmall
        Case "pbLayoutWebPageLarge": PbPublicationLayoutFromString = pbLayoutWebPageLarge
        Case "pbLayoutAdvertisement": PbPublicationLayoutFromString = pbLayoutAdvertisement
        Case "pbLayoutAwardCertificate": PbPublicationLayoutFromString = pbLayoutAwardCertificate
        Case "pbLayoutBanner": PbPublicationLayoutFromString = pbLayoutBanner
        Case "pbLayoutBrochure": PbPublicationLayoutFromString = pbLayoutBrochure
        Case "pbLayoutBusinessCard": PbPublicationLayoutFromString = pbLayoutBusinessCard
        Case "pbLayoutBusinessForm": PbPublicationLayoutFromString = pbLayoutBusinessForm
        Case "pbLayoutCalendar": PbPublicationLayoutFromString = pbLayoutCalendar
        Case "pbLayoutCatalog": PbPublicationLayoutFromString = pbLayoutCatalog
        Case "pbLayoutEmail": PbPublicationLayoutFromString = pbLayoutEmail
        Case "pbLayoutFlyer": PbPublicationLayoutFromString = pbLayoutFlyer
        Case "pbLayoutGiftCertificate": PbPublicationLayoutFromString = pbLayoutGiftCertificate
        Case "pbLayoutGreetingCard": PbPublicationLayoutFromString = pbLayoutGreetingCard
        Case "pbLayoutWordImport": PbPublicationLayoutFromString = pbLayoutWordImport
        Case "pbLayoutInvitationCard": PbPublicationLayoutFromString = pbLayoutInvitationCard
        Case "pbLayoutLetterhead": PbPublicationLayoutFromString = pbLayoutLetterhead
        Case "pbLayoutMenu": PbPublicationLayoutFromString = pbLayoutMenu
        Case "pbLayoutNewsletter": PbPublicationLayoutFromString = pbLayoutNewsletter
        Case "pbLayoutPaperFoldingProject": PbPublicationLayoutFromString = pbLayoutPaperFoldingProject
        Case "pbLayoutPostcard": PbPublicationLayoutFromString = pbLayoutPostcard
        Case "pbLayoutProgram": PbPublicationLayoutFromString = pbLayoutProgram
        Case "pbLayoutResume": PbPublicationLayoutFromString = pbLayoutResume
        Case "pbLayoutSign": PbPublicationLayoutFromString = pbLayoutSign
        Case "pbLayoutWithComplimentsCard": PbPublicationLayoutFromString = pbLayoutWithComplimentsCard
        Case "pbLayoutWebSite": PbPublicationLayoutFromString = pbLayoutWebSite
        Case "pbLayoutQuickPublication": PbPublicationLayoutFromString = pbLayoutQuickPublication
    End Select
End Function

Function PbPublicationLayoutToString(value As PbPublicationLayout) As String
    Select Case value
        Case pbLayoutFullPage: PbPublicationLayoutToString = "pbLayoutFullPage"
        Case pbLayoutBook: PbPublicationLayoutToString = "pbLayoutBook"
        Case pbLayoutFoldCard: PbPublicationLayoutToString = "pbLayoutFoldCard"
        Case pbLayoutGreetingCardL: PbPublicationLayoutToString = "pbLayoutGreetingCardL"
        Case pbLayoutGreetingCardT: PbPublicationLayoutToString = "pbLayoutGreetingCardT"
        Case pbLayoutKookBaePan: PbPublicationLayoutToString = "pbLayoutKookBaePan"
        Case pbLayoutKookPan: PbPublicationLayoutToString = "pbLayoutKookPan"
        Case pbLayoutShinKookPan: PbPublicationLayoutToString = "pbLayoutShinKookPan"
        Case pbLayoutKookBanPan: PbPublicationLayoutToString = "pbLayoutKookBanPan"
        Case pbLayout4x6BaePan: PbPublicationLayoutToString = "pbLayout4x6BaePan"
        Case pbLayout4x6Pan: PbPublicationLayoutToString = "pbLayout4x6Pan"
        Case pbLayout4x6BanPan: PbPublicationLayoutToString = "pbLayout4x6BanPan"
        Case pbLayoutCrownPan: PbPublicationLayoutToString = "pbLayoutCrownPan"
        Case pbLayoutShinSeoPan: PbPublicationLayoutToString = "pbLayoutShinSeoPan"
        Case pbLayoutJang4x6Pan: PbPublicationLayoutToString = "pbLayoutJang4x6Pan"
        Case pbLayoutIndexCard: PbPublicationLayoutToString = "pbLayoutIndexCard"
        Case pbLayoutBusinessCardUS: PbPublicationLayoutToString = "pbLayoutBusinessCardUS"
        Case pbLayoutBusinessCardEurope: PbPublicationLayoutToString = "pbLayoutBusinessCardEurope"
        Case pbLayoutBusinessCardFE: PbPublicationLayoutToString = "pbLayoutBusinessCardFE"
        Case pbLayoutBusinessCardLocal: PbPublicationLayoutToString = "pbLayoutBusinessCardLocal"
        Case pbLayoutPosterSmall: PbPublicationLayoutToString = "pbLayoutPosterSmall"
        Case pbLayoutPosterLarge: PbPublicationLayoutToString = "pbLayoutPosterLarge"
        Case pbLayoutCustom: PbPublicationLayoutToString = "pbLayoutCustom"
        Case pbLayoutBannerSmall: PbPublicationLayoutToString = "pbLayoutBannerSmall"
        Case pbLayoutBannerMedium: PbPublicationLayoutToString = "pbLayoutBannerMedium"
        Case pbLayoutBannerLarge: PbPublicationLayoutToString = "pbLayoutBannerLarge"
        Case pbLayoutBannerCustom: PbPublicationLayoutToString = "pbLayoutBannerCustom"
        Case pbLayoutPostcardUS: PbPublicationLayoutToString = "pbLayoutPostcardUS"
        Case pbLayoutPostcardHalfLetter: PbPublicationLayoutToString = "pbLayoutPostcardHalfLetter"
        Case pbLayoutPostcardA4: PbPublicationLayoutToString = "pbLayoutPostcardA4"
        Case pbLayoutPostcardJapan: PbPublicationLayoutToString = "pbLayoutPostcardJapan"
        Case pbLayoutLabel: PbPublicationLayoutToString = "pbLayoutLabel"
        Case pbLayoutEnvelope: PbPublicationLayoutToString = "pbLayoutEnvelope"
        Case pbLayoutWebPageSmall: PbPublicationLayoutToString = "pbLayoutWebPageSmall"
        Case pbLayoutWebPageLarge: PbPublicationLayoutToString = "pbLayoutWebPageLarge"
        Case pbLayoutAdvertisement: PbPublicationLayoutToString = "pbLayoutAdvertisement"
        Case pbLayoutAwardCertificate: PbPublicationLayoutToString = "pbLayoutAwardCertificate"
        Case pbLayoutBanner: PbPublicationLayoutToString = "pbLayoutBanner"
        Case pbLayoutBrochure: PbPublicationLayoutToString = "pbLayoutBrochure"
        Case pbLayoutBusinessCard: PbPublicationLayoutToString = "pbLayoutBusinessCard"
        Case pbLayoutBusinessForm: PbPublicationLayoutToString = "pbLayoutBusinessForm"
        Case pbLayoutCalendar: PbPublicationLayoutToString = "pbLayoutCalendar"
        Case pbLayoutCatalog: PbPublicationLayoutToString = "pbLayoutCatalog"
        Case pbLayoutEmail: PbPublicationLayoutToString = "pbLayoutEmail"
        Case pbLayoutFlyer: PbPublicationLayoutToString = "pbLayoutFlyer"
        Case pbLayoutGiftCertificate: PbPublicationLayoutToString = "pbLayoutGiftCertificate"
        Case pbLayoutGreetingCard: PbPublicationLayoutToString = "pbLayoutGreetingCard"
        Case pbLayoutWordImport: PbPublicationLayoutToString = "pbLayoutWordImport"
        Case pbLayoutInvitationCard: PbPublicationLayoutToString = "pbLayoutInvitationCard"
        Case pbLayoutLetterhead: PbPublicationLayoutToString = "pbLayoutLetterhead"
        Case pbLayoutMenu: PbPublicationLayoutToString = "pbLayoutMenu"
        Case pbLayoutNewsletter: PbPublicationLayoutToString = "pbLayoutNewsletter"
        Case pbLayoutPaperFoldingProject: PbPublicationLayoutToString = "pbLayoutPaperFoldingProject"
        Case pbLayoutPostcard: PbPublicationLayoutToString = "pbLayoutPostcard"
        Case pbLayoutProgram: PbPublicationLayoutToString = "pbLayoutProgram"
        Case pbLayoutResume: PbPublicationLayoutToString = "pbLayoutResume"
        Case pbLayoutSign: PbPublicationLayoutToString = "pbLayoutSign"
        Case pbLayoutWithComplimentsCard: PbPublicationLayoutToString = "pbLayoutWithComplimentsCard"
        Case pbLayoutWebSite: PbPublicationLayoutToString = "pbLayoutWebSite"
        Case pbLayoutQuickPublication: PbPublicationLayoutToString = "pbLayoutQuickPublication"
    End Select
End Function
