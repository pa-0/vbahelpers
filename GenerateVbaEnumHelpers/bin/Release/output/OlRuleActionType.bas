Attribute VB_Name = "wOlRuleActionType"
Function OlRuleActionTypeFromString(value As String) As OlRuleActionType
    If IsNumeric(value) Then
        OlRuleActionTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olRuleActionUnknown": OlRuleActionTypeFromString = olRuleActionUnknown
        Case "olRuleActionMoveToFolder": OlRuleActionTypeFromString = olRuleActionMoveToFolder
        Case "olRuleActionAssignToCategory": OlRuleActionTypeFromString = olRuleActionAssignToCategory
        Case "olRuleActionDelete": OlRuleActionTypeFromString = olRuleActionDelete
        Case "olRuleActionDeletePermanently": OlRuleActionTypeFromString = olRuleActionDeletePermanently
        Case "olRuleActionCopyToFolder": OlRuleActionTypeFromString = olRuleActionCopyToFolder
        Case "olRuleActionForward": OlRuleActionTypeFromString = olRuleActionForward
        Case "olRuleActionForwardAsAttachment": OlRuleActionTypeFromString = olRuleActionForwardAsAttachment
        Case "olRuleActionRedirect": OlRuleActionTypeFromString = olRuleActionRedirect
        Case "olRuleActionServerReply": OlRuleActionTypeFromString = olRuleActionServerReply
        Case "olRuleActionTemplate": OlRuleActionTypeFromString = olRuleActionTemplate
        Case "olRuleActionFlagForActionInDays": OlRuleActionTypeFromString = olRuleActionFlagForActionInDays
        Case "olRuleActionFlagColor": OlRuleActionTypeFromString = olRuleActionFlagColor
        Case "olRuleActionFlagClear": OlRuleActionTypeFromString = olRuleActionFlagClear
        Case "olRuleActionImportance": OlRuleActionTypeFromString = olRuleActionImportance
        Case "olRuleActionSensitivity": OlRuleActionTypeFromString = olRuleActionSensitivity
        Case "olRuleActionPrint": OlRuleActionTypeFromString = olRuleActionPrint
        Case "olRuleActionPlaySound": OlRuleActionTypeFromString = olRuleActionPlaySound
        Case "olRuleActionStartApplication": OlRuleActionTypeFromString = olRuleActionStartApplication
        Case "olRuleActionMarkRead": OlRuleActionTypeFromString = olRuleActionMarkRead
        Case "olRuleActionRunScript": OlRuleActionTypeFromString = olRuleActionRunScript
        Case "olRuleActionStop": OlRuleActionTypeFromString = olRuleActionStop
        Case "olRuleActionCustomAction": OlRuleActionTypeFromString = olRuleActionCustomAction
        Case "olRuleActionNewItemAlert": OlRuleActionTypeFromString = olRuleActionNewItemAlert
        Case "olRuleActionDesktopAlert": OlRuleActionTypeFromString = olRuleActionDesktopAlert
        Case "olRuleActionNotifyRead": OlRuleActionTypeFromString = olRuleActionNotifyRead
        Case "olRuleActionNotifyDelivery": OlRuleActionTypeFromString = olRuleActionNotifyDelivery
        Case "olRuleActionCcMessage": OlRuleActionTypeFromString = olRuleActionCcMessage
        Case "olRuleActionDefer": OlRuleActionTypeFromString = olRuleActionDefer
        Case "olRuleActionMarkAsTask": OlRuleActionTypeFromString = olRuleActionMarkAsTask
        Case "olRuleActionClearCategories": OlRuleActionTypeFromString = olRuleActionClearCategories
    End Select
End Function

Function OlRuleActionTypeToString(value As OlRuleActionType) As String
    Select Case value
        Case olRuleActionUnknown: OlRuleActionTypeToString = "olRuleActionUnknown"
        Case olRuleActionMoveToFolder: OlRuleActionTypeToString = "olRuleActionMoveToFolder"
        Case olRuleActionAssignToCategory: OlRuleActionTypeToString = "olRuleActionAssignToCategory"
        Case olRuleActionDelete: OlRuleActionTypeToString = "olRuleActionDelete"
        Case olRuleActionDeletePermanently: OlRuleActionTypeToString = "olRuleActionDeletePermanently"
        Case olRuleActionCopyToFolder: OlRuleActionTypeToString = "olRuleActionCopyToFolder"
        Case olRuleActionForward: OlRuleActionTypeToString = "olRuleActionForward"
        Case olRuleActionForwardAsAttachment: OlRuleActionTypeToString = "olRuleActionForwardAsAttachment"
        Case olRuleActionRedirect: OlRuleActionTypeToString = "olRuleActionRedirect"
        Case olRuleActionServerReply: OlRuleActionTypeToString = "olRuleActionServerReply"
        Case olRuleActionTemplate: OlRuleActionTypeToString = "olRuleActionTemplate"
        Case olRuleActionFlagForActionInDays: OlRuleActionTypeToString = "olRuleActionFlagForActionInDays"
        Case olRuleActionFlagColor: OlRuleActionTypeToString = "olRuleActionFlagColor"
        Case olRuleActionFlagClear: OlRuleActionTypeToString = "olRuleActionFlagClear"
        Case olRuleActionImportance: OlRuleActionTypeToString = "olRuleActionImportance"
        Case olRuleActionSensitivity: OlRuleActionTypeToString = "olRuleActionSensitivity"
        Case olRuleActionPrint: OlRuleActionTypeToString = "olRuleActionPrint"
        Case olRuleActionPlaySound: OlRuleActionTypeToString = "olRuleActionPlaySound"
        Case olRuleActionStartApplication: OlRuleActionTypeToString = "olRuleActionStartApplication"
        Case olRuleActionMarkRead: OlRuleActionTypeToString = "olRuleActionMarkRead"
        Case olRuleActionRunScript: OlRuleActionTypeToString = "olRuleActionRunScript"
        Case olRuleActionStop: OlRuleActionTypeToString = "olRuleActionStop"
        Case olRuleActionCustomAction: OlRuleActionTypeToString = "olRuleActionCustomAction"
        Case olRuleActionNewItemAlert: OlRuleActionTypeToString = "olRuleActionNewItemAlert"
        Case olRuleActionDesktopAlert: OlRuleActionTypeToString = "olRuleActionDesktopAlert"
        Case olRuleActionNotifyRead: OlRuleActionTypeToString = "olRuleActionNotifyRead"
        Case olRuleActionNotifyDelivery: OlRuleActionTypeToString = "olRuleActionNotifyDelivery"
        Case olRuleActionCcMessage: OlRuleActionTypeToString = "olRuleActionCcMessage"
        Case olRuleActionDefer: OlRuleActionTypeToString = "olRuleActionDefer"
        Case olRuleActionMarkAsTask: OlRuleActionTypeToString = "olRuleActionMarkAsTask"
        Case olRuleActionClearCategories: OlRuleActionTypeToString = "olRuleActionClearCategories"
    End Select
End Function
