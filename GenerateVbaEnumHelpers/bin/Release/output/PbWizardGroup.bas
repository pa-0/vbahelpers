Attribute VB_Name = "wPbWizardGroup"
Function PbWizardGroupFromString(value As String) As PbWizardGroup
    If IsNumeric(value) Then
        PbWizardGroupFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbWizardGroupLogo": PbWizardGroupFromString = pbWizardGroupLogo
        Case "pbWizardGroupWebCalendars": PbWizardGroupFromString = pbWizardGroupWebCalendars
        Case "pbWizardGroupDots": PbWizardGroupFromString = pbWizardGroupDots
        Case "pbWizardGroupBoxes": PbWizardGroupFromString = pbWizardGroupBoxes
        Case "pbWizardGroupBarbells": PbWizardGroupFromString = pbWizardGroupBarbells
        Case "pbWizardGroupCheckerboards": PbWizardGroupFromString = pbWizardGroupCheckerboards
        Case "pbWizardGroupCoupon": PbWizardGroupFromString = pbWizardGroupCoupon
        Case "pbWizardGroupAttentionGetter": PbWizardGroupFromString = pbWizardGroupAttentionGetter
        Case "pbWizardGroupPhoneTearoff": PbWizardGroupFromString = pbWizardGroupPhoneTearoff
        Case "pbWizardGroupAdvertisements": PbWizardGroupFromString = pbWizardGroupAdvertisements
        Case "pbWizardGroupWebNavigationBars": PbWizardGroupFromString = pbWizardGroupWebNavigationBars
        Case "pbWizardGroupCalendars": PbWizardGroupFromString = pbWizardGroupCalendars
        Case "pbWizardGroupReplyForms": PbWizardGroupFromString = pbWizardGroupReplyForms
        Case "pbWizardGroupJapaneseCalendar": PbWizardGroupFromString = pbWizardGroupJapaneseCalendar
        Case "pbWizardGroupJapaneseAttentionGetters": PbWizardGroupFromString = pbWizardGroupJapaneseAttentionGetters
        Case "pbWizardGroupJapaneseCoupons": PbWizardGroupFromString = pbWizardGroupJapaneseCoupons
        Case "pbWizardGroupWebMastheads": PbWizardGroupFromString = pbWizardGroupWebMastheads
        Case "pbWizardGroupWellPullQuotes": PbWizardGroupFromString = pbWizardGroupWellPullQuotes
        Case "pbWizardGroupWebSidebars": PbWizardGroupFromString = pbWizardGroupWebSidebars
        Case "pbWizardGroupMastheads": PbWizardGroupFromString = pbWizardGroupMastheads
        Case "pbWizardGroupTableOfContents": PbWizardGroupFromString = pbWizardGroupTableOfContents
        Case "pbWizardGroupSidebars": PbWizardGroupFromString = pbWizardGroupSidebars
        Case "pbWizardGroupPullQuotes": PbWizardGroupFromString = pbWizardGroupPullQuotes
        Case "pbWizardGroupPictureCaptions": PbWizardGroupFromString = pbWizardGroupPictureCaptions
        Case "pbWizardGroupWebButtonsEmail": PbWizardGroupFromString = pbWizardGroupWebButtonsEmail
        Case "pbWizardGroupWebButtonsHome": PbWizardGroupFromString = pbWizardGroupWebButtonsHome
        Case "pbWizardGroupWebButtonsLink": PbWizardGroupFromString = pbWizardGroupWebButtonsLink
        Case "pbWizardGroupJapaneseReplyForms": PbWizardGroupFromString = pbWizardGroupJapaneseReplyForms
        Case "pbWizardGroupJapaneseWebMastheads": PbWizardGroupFromString = pbWizardGroupJapaneseWebMastheads
        Case "pbWizardGroupJapaneseWebPullQuotes": PbWizardGroupFromString = pbWizardGroupJapaneseWebPullQuotes
        Case "pbWizardGroupJapaneseWebSidebars": PbWizardGroupFromString = pbWizardGroupJapaneseWebSidebars
        Case "pbWizardGroupJapaneseMastheads": PbWizardGroupFromString = pbWizardGroupJapaneseMastheads
        Case "pbWizardGroupJapaneseTableOfContents": PbWizardGroupFromString = pbWizardGroupJapaneseTableOfContents
        Case "pbWizardGroupJapaneseSidebars": PbWizardGroupFromString = pbWizardGroupJapaneseSidebars
        Case "pbWizardGroupJapanesePullQuotes": PbWizardGroupFromString = pbWizardGroupJapanesePullQuotes
        Case "pbWizardGroupJapaneseWebNavigationBars": PbWizardGroupFromString = pbWizardGroupJapaneseWebNavigationBars
        Case "pbWizardGroupMarquee": PbWizardGroupFromString = pbWizardGroupMarquee
        Case "pbWizardGroupAccentBox": PbWizardGroupFromString = pbWizardGroupAccentBox
        Case "pbWizardGroupPunctuation": PbWizardGroupFromString = pbWizardGroupPunctuation
        Case "pbWizardGroupLinearAccent": PbWizardGroupFromString = pbWizardGroupLinearAccent
        Case "pbWizardGroupAccessoryBar": PbWizardGroupFromString = pbWizardGroupAccessoryBar
        Case "pbWizardGroupBorders": PbWizardGroupFromString = pbWizardGroupBorders
        Case "pbWizardGroupJapaneseMarquees": PbWizardGroupFromString = pbWizardGroupJapaneseMarquees
        Case "pbWizardGroupJapaneseAccentBox": PbWizardGroupFromString = pbWizardGroupJapaneseAccentBox
        Case "pbWizardGroupJapaneseLinearAccent": PbWizardGroupFromString = pbWizardGroupJapaneseLinearAccent
        Case "pbWizardGroupJapaneseAccessoryBar": PbWizardGroupFromString = pbWizardGroupJapaneseAccessoryBar
        Case "pbWizardGroupJapaneseBorders": PbWizardGroupFromString = pbWizardGroupJapaneseBorders
        Case "pbWizardGroupEastAsiaZipCode": PbWizardGroupFromString = pbWizardGroupEastAsiaZipCode
        Case "pbWizardGroupJapaneseWebButtonEmail": PbWizardGroupFromString = pbWizardGroupJapaneseWebButtonEmail
        Case "pbWizardGroupJapaneseWebButtonHome": PbWizardGroupFromString = pbWizardGroupJapaneseWebButtonHome
        Case "pbWizardGroupJapaneseWebButtonLink": PbWizardGroupFromString = pbWizardGroupJapaneseWebButtonLink
    End Select
End Function

Function PbWizardGroupToString(value As PbWizardGroup) As String
    Select Case value
        Case pbWizardGroupLogo: PbWizardGroupToString = "pbWizardGroupLogo"
        Case pbWizardGroupWebCalendars: PbWizardGroupToString = "pbWizardGroupWebCalendars"
        Case pbWizardGroupDots: PbWizardGroupToString = "pbWizardGroupDots"
        Case pbWizardGroupBoxes: PbWizardGroupToString = "pbWizardGroupBoxes"
        Case pbWizardGroupBarbells: PbWizardGroupToString = "pbWizardGroupBarbells"
        Case pbWizardGroupCheckerboards: PbWizardGroupToString = "pbWizardGroupCheckerboards"
        Case pbWizardGroupCoupon: PbWizardGroupToString = "pbWizardGroupCoupon"
        Case pbWizardGroupAttentionGetter: PbWizardGroupToString = "pbWizardGroupAttentionGetter"
        Case pbWizardGroupPhoneTearoff: PbWizardGroupToString = "pbWizardGroupPhoneTearoff"
        Case pbWizardGroupAdvertisements: PbWizardGroupToString = "pbWizardGroupAdvertisements"
        Case pbWizardGroupWebNavigationBars: PbWizardGroupToString = "pbWizardGroupWebNavigationBars"
        Case pbWizardGroupCalendars: PbWizardGroupToString = "pbWizardGroupCalendars"
        Case pbWizardGroupReplyForms: PbWizardGroupToString = "pbWizardGroupReplyForms"
        Case pbWizardGroupJapaneseCalendar: PbWizardGroupToString = "pbWizardGroupJapaneseCalendar"
        Case pbWizardGroupJapaneseAttentionGetters: PbWizardGroupToString = "pbWizardGroupJapaneseAttentionGetters"
        Case pbWizardGroupJapaneseCoupons: PbWizardGroupToString = "pbWizardGroupJapaneseCoupons"
        Case pbWizardGroupWebMastheads: PbWizardGroupToString = "pbWizardGroupWebMastheads"
        Case pbWizardGroupWellPullQuotes: PbWizardGroupToString = "pbWizardGroupWellPullQuotes"
        Case pbWizardGroupWebSidebars: PbWizardGroupToString = "pbWizardGroupWebSidebars"
        Case pbWizardGroupMastheads: PbWizardGroupToString = "pbWizardGroupMastheads"
        Case pbWizardGroupTableOfContents: PbWizardGroupToString = "pbWizardGroupTableOfContents"
        Case pbWizardGroupSidebars: PbWizardGroupToString = "pbWizardGroupSidebars"
        Case pbWizardGroupPullQuotes: PbWizardGroupToString = "pbWizardGroupPullQuotes"
        Case pbWizardGroupPictureCaptions: PbWizardGroupToString = "pbWizardGroupPictureCaptions"
        Case pbWizardGroupWebButtonsEmail: PbWizardGroupToString = "pbWizardGroupWebButtonsEmail"
        Case pbWizardGroupWebButtonsHome: PbWizardGroupToString = "pbWizardGroupWebButtonsHome"
        Case pbWizardGroupWebButtonsLink: PbWizardGroupToString = "pbWizardGroupWebButtonsLink"
        Case pbWizardGroupJapaneseReplyForms: PbWizardGroupToString = "pbWizardGroupJapaneseReplyForms"
        Case pbWizardGroupJapaneseWebMastheads: PbWizardGroupToString = "pbWizardGroupJapaneseWebMastheads"
        Case pbWizardGroupJapaneseWebPullQuotes: PbWizardGroupToString = "pbWizardGroupJapaneseWebPullQuotes"
        Case pbWizardGroupJapaneseWebSidebars: PbWizardGroupToString = "pbWizardGroupJapaneseWebSidebars"
        Case pbWizardGroupJapaneseMastheads: PbWizardGroupToString = "pbWizardGroupJapaneseMastheads"
        Case pbWizardGroupJapaneseTableOfContents: PbWizardGroupToString = "pbWizardGroupJapaneseTableOfContents"
        Case pbWizardGroupJapaneseSidebars: PbWizardGroupToString = "pbWizardGroupJapaneseSidebars"
        Case pbWizardGroupJapanesePullQuotes: PbWizardGroupToString = "pbWizardGroupJapanesePullQuotes"
        Case pbWizardGroupJapaneseWebNavigationBars: PbWizardGroupToString = "pbWizardGroupJapaneseWebNavigationBars"
        Case pbWizardGroupMarquee: PbWizardGroupToString = "pbWizardGroupMarquee"
        Case pbWizardGroupAccentBox: PbWizardGroupToString = "pbWizardGroupAccentBox"
        Case pbWizardGroupPunctuation: PbWizardGroupToString = "pbWizardGroupPunctuation"
        Case pbWizardGroupLinearAccent: PbWizardGroupToString = "pbWizardGroupLinearAccent"
        Case pbWizardGroupAccessoryBar: PbWizardGroupToString = "pbWizardGroupAccessoryBar"
        Case pbWizardGroupBorders: PbWizardGroupToString = "pbWizardGroupBorders"
        Case pbWizardGroupJapaneseMarquees: PbWizardGroupToString = "pbWizardGroupJapaneseMarquees"
        Case pbWizardGroupJapaneseAccentBox: PbWizardGroupToString = "pbWizardGroupJapaneseAccentBox"
        Case pbWizardGroupJapaneseLinearAccent: PbWizardGroupToString = "pbWizardGroupJapaneseLinearAccent"
        Case pbWizardGroupJapaneseAccessoryBar: PbWizardGroupToString = "pbWizardGroupJapaneseAccessoryBar"
        Case pbWizardGroupJapaneseBorders: PbWizardGroupToString = "pbWizardGroupJapaneseBorders"
        Case pbWizardGroupEastAsiaZipCode: PbWizardGroupToString = "pbWizardGroupEastAsiaZipCode"
        Case pbWizardGroupJapaneseWebButtonEmail: PbWizardGroupToString = "pbWizardGroupJapaneseWebButtonEmail"
        Case pbWizardGroupJapaneseWebButtonHome: PbWizardGroupToString = "pbWizardGroupJapaneseWebButtonHome"
        Case pbWizardGroupJapaneseWebButtonLink: PbWizardGroupToString = "pbWizardGroupJapaneseWebButtonLink"
    End Select
End Function
