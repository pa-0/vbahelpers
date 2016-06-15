Attribute VB_Name = "wOlRuleConditionType"
Function OlRuleConditionTypeFromString(value As String) As OlRuleConditionType
    If IsNumeric(value) Then
        OlRuleConditionTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olConditionUnknown": OlRuleConditionTypeFromString = olConditionUnknown
        Case "olConditionFrom": OlRuleConditionTypeFromString = olConditionFrom
        Case "olConditionSubject": OlRuleConditionTypeFromString = olConditionSubject
        Case "olConditionAccount": OlRuleConditionTypeFromString = olConditionAccount
        Case "olConditionOnlyToMe": OlRuleConditionTypeFromString = olConditionOnlyToMe
        Case "olConditionTo": OlRuleConditionTypeFromString = olConditionTo
        Case "olConditionImportance": OlRuleConditionTypeFromString = olConditionImportance
        Case "olConditionSensitivity": OlRuleConditionTypeFromString = olConditionSensitivity
        Case "olConditionFlaggedForAction": OlRuleConditionTypeFromString = olConditionFlaggedForAction
        Case "olConditionCc": OlRuleConditionTypeFromString = olConditionCc
        Case "olConditionToOrCc": OlRuleConditionTypeFromString = olConditionToOrCc
        Case "olConditionNotTo": OlRuleConditionTypeFromString = olConditionNotTo
        Case "olConditionSentTo": OlRuleConditionTypeFromString = olConditionSentTo
        Case "olConditionBody": OlRuleConditionTypeFromString = olConditionBody
        Case "olConditionBodyOrSubject": OlRuleConditionTypeFromString = olConditionBodyOrSubject
        Case "olConditionMessageHeader": OlRuleConditionTypeFromString = olConditionMessageHeader
        Case "olConditionRecipientAddress": OlRuleConditionTypeFromString = olConditionRecipientAddress
        Case "olConditionSenderAddress": OlRuleConditionTypeFromString = olConditionSenderAddress
        Case "olConditionCategory": OlRuleConditionTypeFromString = olConditionCategory
        Case "olConditionOOF": OlRuleConditionTypeFromString = olConditionOOF
        Case "olConditionHasAttachment": OlRuleConditionTypeFromString = olConditionHasAttachment
        Case "olConditionSizeRange": OlRuleConditionTypeFromString = olConditionSizeRange
        Case "olConditionDateRange": OlRuleConditionTypeFromString = olConditionDateRange
        Case "olConditionFormName": OlRuleConditionTypeFromString = olConditionFormName
        Case "olConditionProperty": OlRuleConditionTypeFromString = olConditionProperty
        Case "olConditionSenderInAddressBook": OlRuleConditionTypeFromString = olConditionSenderInAddressBook
        Case "olConditionMeetingInviteOrUpdate": OlRuleConditionTypeFromString = olConditionMeetingInviteOrUpdate
        Case "olConditionLocalMachineOnly": OlRuleConditionTypeFromString = olConditionLocalMachineOnly
        Case "olConditionOtherMachine": OlRuleConditionTypeFromString = olConditionOtherMachine
        Case "olConditionAnyCategory": OlRuleConditionTypeFromString = olConditionAnyCategory
        Case "olConditionFromRssFeed": OlRuleConditionTypeFromString = olConditionFromRssFeed
        Case "olConditionFromAnyRssFeed": OlRuleConditionTypeFromString = olConditionFromAnyRssFeed
    End Select
End Function

Function OlRuleConditionTypeToString(value As OlRuleConditionType) As String
    Select Case value
        Case olConditionUnknown: OlRuleConditionTypeToString = "olConditionUnknown"
        Case olConditionFrom: OlRuleConditionTypeToString = "olConditionFrom"
        Case olConditionSubject: OlRuleConditionTypeToString = "olConditionSubject"
        Case olConditionAccount: OlRuleConditionTypeToString = "olConditionAccount"
        Case olConditionOnlyToMe: OlRuleConditionTypeToString = "olConditionOnlyToMe"
        Case olConditionTo: OlRuleConditionTypeToString = "olConditionTo"
        Case olConditionImportance: OlRuleConditionTypeToString = "olConditionImportance"
        Case olConditionSensitivity: OlRuleConditionTypeToString = "olConditionSensitivity"
        Case olConditionFlaggedForAction: OlRuleConditionTypeToString = "olConditionFlaggedForAction"
        Case olConditionCc: OlRuleConditionTypeToString = "olConditionCc"
        Case olConditionToOrCc: OlRuleConditionTypeToString = "olConditionToOrCc"
        Case olConditionNotTo: OlRuleConditionTypeToString = "olConditionNotTo"
        Case olConditionSentTo: OlRuleConditionTypeToString = "olConditionSentTo"
        Case olConditionBody: OlRuleConditionTypeToString = "olConditionBody"
        Case olConditionBodyOrSubject: OlRuleConditionTypeToString = "olConditionBodyOrSubject"
        Case olConditionMessageHeader: OlRuleConditionTypeToString = "olConditionMessageHeader"
        Case olConditionRecipientAddress: OlRuleConditionTypeToString = "olConditionRecipientAddress"
        Case olConditionSenderAddress: OlRuleConditionTypeToString = "olConditionSenderAddress"
        Case olConditionCategory: OlRuleConditionTypeToString = "olConditionCategory"
        Case olConditionOOF: OlRuleConditionTypeToString = "olConditionOOF"
        Case olConditionHasAttachment: OlRuleConditionTypeToString = "olConditionHasAttachment"
        Case olConditionSizeRange: OlRuleConditionTypeToString = "olConditionSizeRange"
        Case olConditionDateRange: OlRuleConditionTypeToString = "olConditionDateRange"
        Case olConditionFormName: OlRuleConditionTypeToString = "olConditionFormName"
        Case olConditionProperty: OlRuleConditionTypeToString = "olConditionProperty"
        Case olConditionSenderInAddressBook: OlRuleConditionTypeToString = "olConditionSenderInAddressBook"
        Case olConditionMeetingInviteOrUpdate: OlRuleConditionTypeToString = "olConditionMeetingInviteOrUpdate"
        Case olConditionLocalMachineOnly: OlRuleConditionTypeToString = "olConditionLocalMachineOnly"
        Case olConditionOtherMachine: OlRuleConditionTypeToString = "olConditionOtherMachine"
        Case olConditionAnyCategory: OlRuleConditionTypeToString = "olConditionAnyCategory"
        Case olConditionFromRssFeed: OlRuleConditionTypeToString = "olConditionFromRssFeed"
        Case olConditionFromAnyRssFeed: OlRuleConditionTypeToString = "olConditionFromAnyRssFeed"
    End Select
End Function
