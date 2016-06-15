Attribute VB_Name = "wPbWizardPageType10"
Function PbWizardPageType10FromString(value As String) As PbWizardPageType10
    If IsNumeric(value) Then
        PbWizardPageType10FromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbWizardPageTypeWebStory": PbWizardPageType10FromString = pbWizardPageTypeWebStory
        Case "pbWizardPageTypeWebRelatedLinks": PbWizardPageType10FromString = pbWizardPageTypeWebRelatedLinks
        Case "pbWizardPageTypeWebPriceList": PbWizardPageType10FromString = pbWizardPageTypeWebPriceList
        Case "pbWizardPageTypeWebCalendar": PbWizardPageType10FromString = pbWizardPageTypeWebCalendar
        Case "pbWizardPageTypeWebSpecialOffer": PbWizardPageType10FromString = pbWizardPageTypeWebSpecialOffer
        Case "pbWizardPageTypeWebEvent": PbWizardPageType10FromString = pbWizardPageTypeWebEvent
    End Select
End Function

Function PbWizardPageType10ToString(value As PbWizardPageType10) As String
    Select Case value
        Case pbWizardPageTypeWebStory: PbWizardPageType10ToString = "pbWizardPageTypeWebStory"
        Case pbWizardPageTypeWebRelatedLinks: PbWizardPageType10ToString = "pbWizardPageTypeWebRelatedLinks"
        Case pbWizardPageTypeWebPriceList: PbWizardPageType10ToString = "pbWizardPageTypeWebPriceList"
        Case pbWizardPageTypeWebCalendar: PbWizardPageType10ToString = "pbWizardPageTypeWebCalendar"
        Case pbWizardPageTypeWebSpecialOffer: PbWizardPageType10ToString = "pbWizardPageTypeWebSpecialOffer"
        Case pbWizardPageTypeWebEvent: PbWizardPageType10ToString = "pbWizardPageTypeWebEvent"
    End Select
End Function
