Attribute VB_Name = "wPbWizardPageType"
Function PbWizardPageTypeFromString(value As String) As PbWizardPageType
    If IsNumeric(value) Then
        PbWizardPageTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbWizardPageTypeNewsletter3Stories": PbWizardPageTypeFromString = pbWizardPageTypeNewsletter3Stories
        Case "pbWizardPageTypeNewsletterCalendar": PbWizardPageTypeFromString = pbWizardPageTypeNewsletterCalendar
        Case "pbWizardPageTypeNewsletterOrderForm": PbWizardPageTypeFromString = pbWizardPageTypeNewsletterOrderForm
        Case "pbWizardPageTypeNewsletterResponseForm": PbWizardPageTypeFromString = pbWizardPageTypeNewsletterResponseForm
        Case "pbWizardPageTypeNewsletterSignupForm": PbWizardPageTypeFromString = pbWizardPageTypeNewsletterSignupForm
        Case "pbWizardPageTypeCatalogOneColumnText": PbWizardPageTypeFromString = pbWizardPageTypeCatalogOneColumnText
        Case "pbWizardPageTypeCatalogOneColumnTextPicture": PbWizardPageTypeFromString = pbWizardPageTypeCatalogOneColumnTextPicture
        Case "pbWizardPageTypeCatalogTwoColumnsText": PbWizardPageTypeFromString = pbWizardPageTypeCatalogTwoColumnsText
        Case "pbWizardPageTypeCatalogTwoColumnsTextPicture": PbWizardPageTypeFromString = pbWizardPageTypeCatalogTwoColumnsTextPicture
        Case "pbWizardPageTypeCatalogCalendar": PbWizardPageTypeFromString = pbWizardPageTypeCatalogCalendar
        Case "pbWizardPageTypeCatalogTableOfContents": PbWizardPageTypeFromString = pbWizardPageTypeCatalogTableOfContents
        Case "pbWizardPageTypeCatalogFeaturedItem": PbWizardPageTypeFromString = pbWizardPageTypeCatalogFeaturedItem
        Case "pbWizardPageTypeCatalogTwoItemsAlignedPictures": PbWizardPageTypeFromString = pbWizardPageTypeCatalogTwoItemsAlignedPictures
        Case "pbWizardPageTypeCatalogTwoItemsOffsetPictures": PbWizardPageTypeFromString = pbWizardPageTypeCatalogTwoItemsOffsetPictures
        Case "pbWizardPageTypeCatalogThreeItemsAlignedPictures": PbWizardPageTypeFromString = pbWizardPageTypeCatalogThreeItemsAlignedPictures
        Case "pbWizardPageTypeCatalogThreeItemsOffsetPictures": PbWizardPageTypeFromString = pbWizardPageTypeCatalogThreeItemsOffsetPictures
        Case "pbWizardPageTypeCatalogThreeItemsStackedPictures": PbWizardPageTypeFromString = pbWizardPageTypeCatalogThreeItemsStackedPictures
        Case "pbWizardPageTypeCatalogFourItemsAlignedPictures": PbWizardPageTypeFromString = pbWizardPageTypeCatalogFourItemsAlignedPictures
        Case "pbWizardPageTypeCatalogFourItemsOffsetPictures": PbWizardPageTypeFromString = pbWizardPageTypeCatalogFourItemsOffsetPictures
        Case "pbWizardPageTypeCatalogFourItemsSquaredPictures": PbWizardPageTypeFromString = pbWizardPageTypeCatalogFourItemsSquaredPictures
        Case "pbWizardPageTypeCatalogEightItemsOneColumn": PbWizardPageTypeFromString = pbWizardPageTypeCatalogEightItemsOneColumn
        Case "pbWizardPageTypeCatalogEightItemsTwoColumns": PbWizardPageTypeFromString = pbWizardPageTypeCatalogEightItemsTwoColumns
        Case "pbWizardPageTypeCatalogBlank": PbWizardPageTypeFromString = pbWizardPageTypeCatalogBlank
        Case "pbWizardPageTypeCatalogForm": PbWizardPageTypeFromString = pbWizardPageTypeCatalogForm
        Case "pbWizardPageTypeWebAboutUs": PbWizardPageTypeFromString = pbWizardPageTypeWebAboutUs
        Case "pbWizardPageTypeWebInformational": PbWizardPageTypeFromString = pbWizardPageTypeWebInformational
        Case "pbWizardPageTypeWebList": PbWizardPageTypeFromString = pbWizardPageTypeWebList
        Case "pbWizardPageTypeWebCalendarPage": PbWizardPageTypeFromString = pbWizardPageTypeWebCalendarPage
        Case "pbWizardPageTypeWebContactUs": PbWizardPageTypeFromString = pbWizardPageTypeWebContactUs
        Case "pbWizardPageTypeWebEmployeeList": PbWizardPageTypeFromString = pbWizardPageTypeWebEmployeeList
        Case "pbWizardPageTypeWebEmployee": PbWizardPageTypeFromString = pbWizardPageTypeWebEmployee
        Case "pbWizardPageTypeWebFAQ": PbWizardPageTypeFromString = pbWizardPageTypeWebFAQ
        Case "pbWizardPageTypeWebHome": PbWizardPageTypeFromString = pbWizardPageTypeWebHome
        Case "pbWizardPageTypeWebJobs": PbWizardPageTypeFromString = pbWizardPageTypeWebJobs
        Case "pbWizardPageTypeWebLegal": PbWizardPageTypeFromString = pbWizardPageTypeWebLegal
        Case "pbWizardPageTypeWebArticle": PbWizardPageTypeFromString = pbWizardPageTypeWebArticle
        Case "pbWizardPageTypeWebPhoto": PbWizardPageTypeFromString = pbWizardPageTypeWebPhoto
        Case "pbWizardPageTypeWebPhotoGallery": PbWizardPageTypeFromString = pbWizardPageTypeWebPhotoGallery
        Case "pbWizardPageTypeWebProduct": PbWizardPageTypeFromString = pbWizardPageTypeWebProduct
        Case "pbWizardPageTypeWebProductList": PbWizardPageTypeFromString = pbWizardPageTypeWebProductList
        Case "pbWizardPageTypeWebProjectList": PbWizardPageTypeFromString = pbWizardPageTypeWebProjectList
        Case "pbWizardPageTypeWebLinks": PbWizardPageTypeFromString = pbWizardPageTypeWebLinks
        Case "pbWizardPageTypeWebSeminar": PbWizardPageTypeFromString = pbWizardPageTypeWebSeminar
        Case "pbWizardPageTypeWebServiceList": PbWizardPageTypeFromString = pbWizardPageTypeWebServiceList
        Case "pbWizardPageTypeWebService": PbWizardPageTypeFromString = pbWizardPageTypeWebService
        Case "pbWizardPageTypeWebSpecial": PbWizardPageTypeFromString = pbWizardPageTypeWebSpecial
        Case "pbWizardPageTypeWebBlank": PbWizardPageTypeFromString = pbWizardPageTypeWebBlank
        Case "pbWizardPageTypeWebOrderForm": PbWizardPageTypeFromString = pbWizardPageTypeWebOrderForm
        Case "pbWizardPageTypeWebResponseForm": PbWizardPageTypeFromString = pbWizardPageTypeWebResponseForm
        Case "pbWizardPageTypeWebSignupForm": PbWizardPageTypeFromString = pbWizardPageTypeWebSignupForm
        Case "pbWizardPageTypeWebCalendarWithLinks": PbWizardPageTypeFromString = pbWizardPageTypeWebCalendarWithLinks
        Case "pbWizardPageTypeWebProductsWithLinks": PbWizardPageTypeFromString = pbWizardPageTypeWebProductsWithLinks
        Case "pbWizardPageTypeWebEmployeesWithLinks": PbWizardPageTypeFromString = pbWizardPageTypeWebEmployeesWithLinks
        Case "pbWizardPageTypeWebServicesWithLinks": PbWizardPageTypeFromString = pbWizardPageTypeWebServicesWithLinks
        Case "pbWizardPageTypeWebProjectsWithLinks": PbWizardPageTypeFromString = pbWizardPageTypeWebProjectsWithLinks
        Case "pbWizardPageTypeWebPhotosWithLinks": PbWizardPageTypeFromString = pbWizardPageTypeWebPhotosWithLinks
        Case "pbWizardPageTypeNone": PbWizardPageTypeFromString = pbWizardPageTypeNone
    End Select
End Function

Function PbWizardPageTypeToString(value As PbWizardPageType) As String
    Select Case value
        Case pbWizardPageTypeNewsletter3Stories: PbWizardPageTypeToString = "pbWizardPageTypeNewsletter3Stories"
        Case pbWizardPageTypeNewsletterCalendar: PbWizardPageTypeToString = "pbWizardPageTypeNewsletterCalendar"
        Case pbWizardPageTypeNewsletterOrderForm: PbWizardPageTypeToString = "pbWizardPageTypeNewsletterOrderForm"
        Case pbWizardPageTypeNewsletterResponseForm: PbWizardPageTypeToString = "pbWizardPageTypeNewsletterResponseForm"
        Case pbWizardPageTypeNewsletterSignupForm: PbWizardPageTypeToString = "pbWizardPageTypeNewsletterSignupForm"
        Case pbWizardPageTypeCatalogOneColumnText: PbWizardPageTypeToString = "pbWizardPageTypeCatalogOneColumnText"
        Case pbWizardPageTypeCatalogOneColumnTextPicture: PbWizardPageTypeToString = "pbWizardPageTypeCatalogOneColumnTextPicture"
        Case pbWizardPageTypeCatalogTwoColumnsText: PbWizardPageTypeToString = "pbWizardPageTypeCatalogTwoColumnsText"
        Case pbWizardPageTypeCatalogTwoColumnsTextPicture: PbWizardPageTypeToString = "pbWizardPageTypeCatalogTwoColumnsTextPicture"
        Case pbWizardPageTypeCatalogCalendar: PbWizardPageTypeToString = "pbWizardPageTypeCatalogCalendar"
        Case pbWizardPageTypeCatalogTableOfContents: PbWizardPageTypeToString = "pbWizardPageTypeCatalogTableOfContents"
        Case pbWizardPageTypeCatalogFeaturedItem: PbWizardPageTypeToString = "pbWizardPageTypeCatalogFeaturedItem"
        Case pbWizardPageTypeCatalogTwoItemsAlignedPictures: PbWizardPageTypeToString = "pbWizardPageTypeCatalogTwoItemsAlignedPictures"
        Case pbWizardPageTypeCatalogTwoItemsOffsetPictures: PbWizardPageTypeToString = "pbWizardPageTypeCatalogTwoItemsOffsetPictures"
        Case pbWizardPageTypeCatalogThreeItemsAlignedPictures: PbWizardPageTypeToString = "pbWizardPageTypeCatalogThreeItemsAlignedPictures"
        Case pbWizardPageTypeCatalogThreeItemsOffsetPictures: PbWizardPageTypeToString = "pbWizardPageTypeCatalogThreeItemsOffsetPictures"
        Case pbWizardPageTypeCatalogThreeItemsStackedPictures: PbWizardPageTypeToString = "pbWizardPageTypeCatalogThreeItemsStackedPictures"
        Case pbWizardPageTypeCatalogFourItemsAlignedPictures: PbWizardPageTypeToString = "pbWizardPageTypeCatalogFourItemsAlignedPictures"
        Case pbWizardPageTypeCatalogFourItemsOffsetPictures: PbWizardPageTypeToString = "pbWizardPageTypeCatalogFourItemsOffsetPictures"
        Case pbWizardPageTypeCatalogFourItemsSquaredPictures: PbWizardPageTypeToString = "pbWizardPageTypeCatalogFourItemsSquaredPictures"
        Case pbWizardPageTypeCatalogEightItemsOneColumn: PbWizardPageTypeToString = "pbWizardPageTypeCatalogEightItemsOneColumn"
        Case pbWizardPageTypeCatalogEightItemsTwoColumns: PbWizardPageTypeToString = "pbWizardPageTypeCatalogEightItemsTwoColumns"
        Case pbWizardPageTypeCatalogBlank: PbWizardPageTypeToString = "pbWizardPageTypeCatalogBlank"
        Case pbWizardPageTypeCatalogForm: PbWizardPageTypeToString = "pbWizardPageTypeCatalogForm"
        Case pbWizardPageTypeWebAboutUs: PbWizardPageTypeToString = "pbWizardPageTypeWebAboutUs"
        Case pbWizardPageTypeWebInformational: PbWizardPageTypeToString = "pbWizardPageTypeWebInformational"
        Case pbWizardPageTypeWebList: PbWizardPageTypeToString = "pbWizardPageTypeWebList"
        Case pbWizardPageTypeWebCalendarPage: PbWizardPageTypeToString = "pbWizardPageTypeWebCalendarPage"
        Case pbWizardPageTypeWebContactUs: PbWizardPageTypeToString = "pbWizardPageTypeWebContactUs"
        Case pbWizardPageTypeWebEmployeeList: PbWizardPageTypeToString = "pbWizardPageTypeWebEmployeeList"
        Case pbWizardPageTypeWebEmployee: PbWizardPageTypeToString = "pbWizardPageTypeWebEmployee"
        Case pbWizardPageTypeWebFAQ: PbWizardPageTypeToString = "pbWizardPageTypeWebFAQ"
        Case pbWizardPageTypeWebHome: PbWizardPageTypeToString = "pbWizardPageTypeWebHome"
        Case pbWizardPageTypeWebJobs: PbWizardPageTypeToString = "pbWizardPageTypeWebJobs"
        Case pbWizardPageTypeWebLegal: PbWizardPageTypeToString = "pbWizardPageTypeWebLegal"
        Case pbWizardPageTypeWebArticle: PbWizardPageTypeToString = "pbWizardPageTypeWebArticle"
        Case pbWizardPageTypeWebPhoto: PbWizardPageTypeToString = "pbWizardPageTypeWebPhoto"
        Case pbWizardPageTypeWebPhotoGallery: PbWizardPageTypeToString = "pbWizardPageTypeWebPhotoGallery"
        Case pbWizardPageTypeWebProduct: PbWizardPageTypeToString = "pbWizardPageTypeWebProduct"
        Case pbWizardPageTypeWebProductList: PbWizardPageTypeToString = "pbWizardPageTypeWebProductList"
        Case pbWizardPageTypeWebProjectList: PbWizardPageTypeToString = "pbWizardPageTypeWebProjectList"
        Case pbWizardPageTypeWebLinks: PbWizardPageTypeToString = "pbWizardPageTypeWebLinks"
        Case pbWizardPageTypeWebSeminar: PbWizardPageTypeToString = "pbWizardPageTypeWebSeminar"
        Case pbWizardPageTypeWebServiceList: PbWizardPageTypeToString = "pbWizardPageTypeWebServiceList"
        Case pbWizardPageTypeWebService: PbWizardPageTypeToString = "pbWizardPageTypeWebService"
        Case pbWizardPageTypeWebSpecial: PbWizardPageTypeToString = "pbWizardPageTypeWebSpecial"
        Case pbWizardPageTypeWebBlank: PbWizardPageTypeToString = "pbWizardPageTypeWebBlank"
        Case pbWizardPageTypeWebOrderForm: PbWizardPageTypeToString = "pbWizardPageTypeWebOrderForm"
        Case pbWizardPageTypeWebResponseForm: PbWizardPageTypeToString = "pbWizardPageTypeWebResponseForm"
        Case pbWizardPageTypeWebSignupForm: PbWizardPageTypeToString = "pbWizardPageTypeWebSignupForm"
        Case pbWizardPageTypeWebCalendarWithLinks: PbWizardPageTypeToString = "pbWizardPageTypeWebCalendarWithLinks"
        Case pbWizardPageTypeWebProductsWithLinks: PbWizardPageTypeToString = "pbWizardPageTypeWebProductsWithLinks"
        Case pbWizardPageTypeWebEmployeesWithLinks: PbWizardPageTypeToString = "pbWizardPageTypeWebEmployeesWithLinks"
        Case pbWizardPageTypeWebServicesWithLinks: PbWizardPageTypeToString = "pbWizardPageTypeWebServicesWithLinks"
        Case pbWizardPageTypeWebProjectsWithLinks: PbWizardPageTypeToString = "pbWizardPageTypeWebProjectsWithLinks"
        Case pbWizardPageTypeWebPhotosWithLinks: PbWizardPageTypeToString = "pbWizardPageTypeWebPhotosWithLinks"
        Case pbWizardPageTypeNone: PbWizardPageTypeToString = "pbWizardPageTypeNone"
    End Select
End Function
