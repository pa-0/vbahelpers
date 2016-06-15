Attribute VB_Name = "wPbWizardTag"
Function PbWizardTagFromString(value As String) As PbWizardTag
    If IsNumeric(value) Then
        PbWizardTagFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbWizardTagLogoGroup": PbWizardTagFromString = pbWizardTagLogoGroup
        Case "pbWizardTagOrganizationName": PbWizardTagFromString = pbWizardTagOrganizationName
        Case "pbWizardTagPersonalName": PbWizardTagFromString = pbWizardTagPersonalName
        Case "pbWizardTagAddress": PbWizardTagFromString = pbWizardTagAddress
        Case "pbWizardTagTagLine": PbWizardTagFromString = pbWizardTagTagLine
        Case "pbWizardTagPhoneFaxEmail": PbWizardTagFromString = pbWizardTagPhoneFaxEmail
        Case "pbWizardTagPhoneNumber": PbWizardTagFromString = pbWizardTagPhoneNumber
        Case "pbWizardTagJobTitle": PbWizardTagFromString = pbWizardTagJobTitle
        Case "pbWizardTagPersonalNameGroup": PbWizardTagFromString = pbWizardTagPersonalNameGroup
        Case "pbWizardTagAddressGroup": PbWizardTagFromString = pbWizardTagAddressGroup
        Case "pbWizardTagOrganizationNameGroup": PbWizardTagFromString = pbWizardTagOrganizationNameGroup
        Case "pbWizardTagTagLineGroup": PbWizardTagFromString = pbWizardTagTagLineGroup
        Case "pbWizardTagPhoneFaxEmailGroup": PbWizardTagFromString = pbWizardTagPhoneFaxEmailGroup
        Case "pbWizardTagLocation": PbWizardTagFromString = pbWizardTagLocation
        Case "pbWizardTagMapPicture": PbWizardTagFromString = pbWizardTagMapPicture
        Case "pbWizardTagCustomerMailingAddress": PbWizardTagFromString = pbWizardTagCustomerMailingAddress
        Case "pbWizardTagHourTimeDateInformation": PbWizardTagFromString = pbWizardTagHourTimeDateInformation
        Case "pbWizardTagBusinessDescription": PbWizardTagFromString = pbWizardTagBusinessDescription
        Case "pbWizardTagReturnAddressLines": PbWizardTagFromString = pbWizardTagReturnAddressLines
        Case "pbWizardTagStampBoxOutline": PbWizardTagFromString = pbWizardTagStampBoxOutline
        Case "pbWizardTagStampBox": PbWizardTagFromString = pbWizardTagStampBox
        Case "pbWizardTagPhotoPlaceholderFrame": PbWizardTagFromString = pbWizardTagPhotoPlaceholderFrame
        Case "pbWizardTagPhotePlaceholderText": PbWizardTagFromString = pbWizardTagPhotePlaceholderText
        Case "pbWizardTagPublicationDate": PbWizardTagFromString = pbWizardTagPublicationDate
        Case "pbWizardTagTableOfContentsTitle": PbWizardTagFromString = pbWizardTagTableOfContentsTitle
        Case "pbWizardTagTableOfContents": PbWizardTagFromString = pbWizardTagTableOfContents
        Case "pbWizardTagNewsletterTitle": PbWizardTagFromString = pbWizardTagNewsletterTitle
        Case "pbWizardTagPageNumber": PbWizardTagFromString = pbWizardTagPageNumber
        Case "pbWizardTagStoryTitle": PbWizardTagFromString = pbWizardTagStoryTitle
        Case "pbWizardTagStory": PbWizardTagFromString = pbWizardTagStory
        Case "pbWizardTagStoryGraphicPrimary": PbWizardTagFromString = pbWizardTagStoryGraphicPrimary
        Case "pbWizardTagStoryCaptionPrimary": PbWizardTagFromString = pbWizardTagStoryCaptionPrimary
        Case "pbWizardTagBriefDescriptionSummary": PbWizardTagFromString = pbWizardTagBriefDescriptionSummary
        Case "pbWizardTagLinkedStoryPrimary": PbWizardTagFromString = pbWizardTagLinkedStoryPrimary
        Case "pbWizardTagLinkedStorySecondary": PbWizardTagFromString = pbWizardTagLinkedStorySecondary
        Case "pbWizardTagLinkedStoryTertiary": PbWizardTagFromString = pbWizardTagLinkedStoryTertiary
        Case "pbWizardTagMainFloatingGraphic": PbWizardTagFromString = pbWizardTagMainFloatingGraphic
        Case "pbWizardTagBriefDescriptionGraphic": PbWizardTagFromString = pbWizardTagBriefDescriptionGraphic
        Case "pbWizardTagStoryGraphicSecondary": PbWizardTagFromString = pbWizardTagStoryGraphicSecondary
        Case "pbWizardTagBriefDescriptionCaption": PbWizardTagFromString = pbWizardTagBriefDescriptionCaption
        Case "pbWizardTagFloatingGraphicCaption": PbWizardTagFromString = pbWizardTagFloatingGraphicCaption
        Case "pbWizardTagBriefDescriptionTitle": PbWizardTagFromString = pbWizardTagBriefDescriptionTitle
        Case "pbWizardTagBriefDescriptionSummaryPrimary": PbWizardTagFromString = pbWizardTagBriefDescriptionSummaryPrimary
        Case "pbWizardTagStoryCaptionSecondary": PbWizardTagFromString = pbWizardTagStoryCaptionSecondary
        Case "pbWizardTagMasthead": PbWizardTagFromString = pbWizardTagMasthead
        Case "pbWizardTagMainTitle": PbWizardTagFromString = pbWizardTagMainTitle
        Case "pbWizardTagMainGraphic": PbWizardTagFromString = pbWizardTagMainGraphic
        Case "pbWizardTagDate": PbWizardTagFromString = pbWizardTagDate
        Case "pbWizardTagTime": PbWizardTagFromString = pbWizardTagTime
        Case "pbWizardTagList": PbWizardTagFromString = pbWizardTagList
        Case "pbWizardTagQuickPubHeading": PbWizardTagFromString = pbWizardTagQuickPubHeading
        Case "pbWizardTagQuickPubMessage": PbWizardTagFromString = pbWizardTagQuickPubMessage
        Case "pbWizardTagQuickPubPicture": PbWizardTagFromString = pbWizardTagQuickPubPicture
        Case "pbWizardTagQuickPubContent": PbWizardTagFromString = pbWizardTagQuickPubContent
        Case "pbWizardTagEAPostalCodeGroup": PbWizardTagFromString = pbWizardTagEAPostalCodeGroup
        Case "pbWizardTagEAPostalCodeBox": PbWizardTagFromString = pbWizardTagEAPostalCodeBox
        Case "pbWizardTagEAPostalCodeLine": PbWizardTagFromString = pbWizardTagEAPostalCodeLine
        Case "pbWizardTagPictureCaptionPicture": PbWizardTagFromString = pbWizardTagPictureCaptionPicture
        Case "pbWizardTagPictureCaptionText": PbWizardTagFromString = pbWizardTagPictureCaptionText
        Case "pbWizardTagPictureCaption": PbWizardTagFromString = pbWizardTagPictureCaption
    End Select
End Function

Function PbWizardTagToString(value As PbWizardTag) As String
    Select Case value
        Case pbWizardTagLogoGroup: PbWizardTagToString = "pbWizardTagLogoGroup"
        Case pbWizardTagOrganizationName: PbWizardTagToString = "pbWizardTagOrganizationName"
        Case pbWizardTagPersonalName: PbWizardTagToString = "pbWizardTagPersonalName"
        Case pbWizardTagAddress: PbWizardTagToString = "pbWizardTagAddress"
        Case pbWizardTagTagLine: PbWizardTagToString = "pbWizardTagTagLine"
        Case pbWizardTagPhoneFaxEmail: PbWizardTagToString = "pbWizardTagPhoneFaxEmail"
        Case pbWizardTagPhoneNumber: PbWizardTagToString = "pbWizardTagPhoneNumber"
        Case pbWizardTagJobTitle: PbWizardTagToString = "pbWizardTagJobTitle"
        Case pbWizardTagPersonalNameGroup: PbWizardTagToString = "pbWizardTagPersonalNameGroup"
        Case pbWizardTagAddressGroup: PbWizardTagToString = "pbWizardTagAddressGroup"
        Case pbWizardTagOrganizationNameGroup: PbWizardTagToString = "pbWizardTagOrganizationNameGroup"
        Case pbWizardTagTagLineGroup: PbWizardTagToString = "pbWizardTagTagLineGroup"
        Case pbWizardTagPhoneFaxEmailGroup: PbWizardTagToString = "pbWizardTagPhoneFaxEmailGroup"
        Case pbWizardTagLocation: PbWizardTagToString = "pbWizardTagLocation"
        Case pbWizardTagMapPicture: PbWizardTagToString = "pbWizardTagMapPicture"
        Case pbWizardTagCustomerMailingAddress: PbWizardTagToString = "pbWizardTagCustomerMailingAddress"
        Case pbWizardTagHourTimeDateInformation: PbWizardTagToString = "pbWizardTagHourTimeDateInformation"
        Case pbWizardTagBusinessDescription: PbWizardTagToString = "pbWizardTagBusinessDescription"
        Case pbWizardTagReturnAddressLines: PbWizardTagToString = "pbWizardTagReturnAddressLines"
        Case pbWizardTagStampBoxOutline: PbWizardTagToString = "pbWizardTagStampBoxOutline"
        Case pbWizardTagStampBox: PbWizardTagToString = "pbWizardTagStampBox"
        Case pbWizardTagPhotoPlaceholderFrame: PbWizardTagToString = "pbWizardTagPhotoPlaceholderFrame"
        Case pbWizardTagPhotePlaceholderText: PbWizardTagToString = "pbWizardTagPhotePlaceholderText"
        Case pbWizardTagPublicationDate: PbWizardTagToString = "pbWizardTagPublicationDate"
        Case pbWizardTagTableOfContentsTitle: PbWizardTagToString = "pbWizardTagTableOfContentsTitle"
        Case pbWizardTagTableOfContents: PbWizardTagToString = "pbWizardTagTableOfContents"
        Case pbWizardTagNewsletterTitle: PbWizardTagToString = "pbWizardTagNewsletterTitle"
        Case pbWizardTagPageNumber: PbWizardTagToString = "pbWizardTagPageNumber"
        Case pbWizardTagStoryTitle: PbWizardTagToString = "pbWizardTagStoryTitle"
        Case pbWizardTagStory: PbWizardTagToString = "pbWizardTagStory"
        Case pbWizardTagStoryGraphicPrimary: PbWizardTagToString = "pbWizardTagStoryGraphicPrimary"
        Case pbWizardTagStoryCaptionPrimary: PbWizardTagToString = "pbWizardTagStoryCaptionPrimary"
        Case pbWizardTagBriefDescriptionSummary: PbWizardTagToString = "pbWizardTagBriefDescriptionSummary"
        Case pbWizardTagLinkedStoryPrimary: PbWizardTagToString = "pbWizardTagLinkedStoryPrimary"
        Case pbWizardTagLinkedStorySecondary: PbWizardTagToString = "pbWizardTagLinkedStorySecondary"
        Case pbWizardTagLinkedStoryTertiary: PbWizardTagToString = "pbWizardTagLinkedStoryTertiary"
        Case pbWizardTagMainFloatingGraphic: PbWizardTagToString = "pbWizardTagMainFloatingGraphic"
        Case pbWizardTagBriefDescriptionGraphic: PbWizardTagToString = "pbWizardTagBriefDescriptionGraphic"
        Case pbWizardTagStoryGraphicSecondary: PbWizardTagToString = "pbWizardTagStoryGraphicSecondary"
        Case pbWizardTagBriefDescriptionCaption: PbWizardTagToString = "pbWizardTagBriefDescriptionCaption"
        Case pbWizardTagFloatingGraphicCaption: PbWizardTagToString = "pbWizardTagFloatingGraphicCaption"
        Case pbWizardTagBriefDescriptionTitle: PbWizardTagToString = "pbWizardTagBriefDescriptionTitle"
        Case pbWizardTagBriefDescriptionSummaryPrimary: PbWizardTagToString = "pbWizardTagBriefDescriptionSummaryPrimary"
        Case pbWizardTagStoryCaptionSecondary: PbWizardTagToString = "pbWizardTagStoryCaptionSecondary"
        Case pbWizardTagMasthead: PbWizardTagToString = "pbWizardTagMasthead"
        Case pbWizardTagMainTitle: PbWizardTagToString = "pbWizardTagMainTitle"
        Case pbWizardTagMainGraphic: PbWizardTagToString = "pbWizardTagMainGraphic"
        Case pbWizardTagDate: PbWizardTagToString = "pbWizardTagDate"
        Case pbWizardTagTime: PbWizardTagToString = "pbWizardTagTime"
        Case pbWizardTagList: PbWizardTagToString = "pbWizardTagList"
        Case pbWizardTagQuickPubHeading: PbWizardTagToString = "pbWizardTagQuickPubHeading"
        Case pbWizardTagQuickPubMessage: PbWizardTagToString = "pbWizardTagQuickPubMessage"
        Case pbWizardTagQuickPubPicture: PbWizardTagToString = "pbWizardTagQuickPubPicture"
        Case pbWizardTagQuickPubContent: PbWizardTagToString = "pbWizardTagQuickPubContent"
        Case pbWizardTagEAPostalCodeGroup: PbWizardTagToString = "pbWizardTagEAPostalCodeGroup"
        Case pbWizardTagEAPostalCodeBox: PbWizardTagToString = "pbWizardTagEAPostalCodeBox"
        Case pbWizardTagEAPostalCodeLine: PbWizardTagToString = "pbWizardTagEAPostalCodeLine"
        Case pbWizardTagPictureCaptionPicture: PbWizardTagToString = "pbWizardTagPictureCaptionPicture"
        Case pbWizardTagPictureCaptionText: PbWizardTagToString = "pbWizardTagPictureCaptionText"
        Case pbWizardTagPictureCaption: PbWizardTagToString = "pbWizardTagPictureCaption"
    End Select
End Function
