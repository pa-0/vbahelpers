Attribute VB_Name = "wWdSmartTagControlType"
Function WdSmartTagControlTypeFromString(value As String) As WdSmartTagControlType
    If IsNumeric(value) Then
        WdSmartTagControlTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdControlSmartTag": WdSmartTagControlTypeFromString = wdControlSmartTag
        Case "wdControlLink": WdSmartTagControlTypeFromString = wdControlLink
        Case "wdControlHelp": WdSmartTagControlTypeFromString = wdControlHelp
        Case "wdControlHelpURL": WdSmartTagControlTypeFromString = wdControlHelpURL
        Case "wdControlSeparator": WdSmartTagControlTypeFromString = wdControlSeparator
        Case "wdControlButton": WdSmartTagControlTypeFromString = wdControlButton
        Case "wdControlLabel": WdSmartTagControlTypeFromString = wdControlLabel
        Case "wdControlImage": WdSmartTagControlTypeFromString = wdControlImage
        Case "wdControlCheckbox": WdSmartTagControlTypeFromString = wdControlCheckbox
        Case "wdControlTextbox": WdSmartTagControlTypeFromString = wdControlTextbox
        Case "wdControlListbox": WdSmartTagControlTypeFromString = wdControlListbox
        Case "wdControlCombo": WdSmartTagControlTypeFromString = wdControlCombo
        Case "wdControlActiveX": WdSmartTagControlTypeFromString = wdControlActiveX
        Case "wdControlDocumentFragment": WdSmartTagControlTypeFromString = wdControlDocumentFragment
        Case "wdControlDocumentFragmentURL": WdSmartTagControlTypeFromString = wdControlDocumentFragmentURL
        Case "wdControlRadioGroup": WdSmartTagControlTypeFromString = wdControlRadioGroup
    End Select
End Function

Function WdSmartTagControlTypeToString(value As WdSmartTagControlType) As String
    Select Case value
        Case wdControlSmartTag: WdSmartTagControlTypeToString = "wdControlSmartTag"
        Case wdControlLink: WdSmartTagControlTypeToString = "wdControlLink"
        Case wdControlHelp: WdSmartTagControlTypeToString = "wdControlHelp"
        Case wdControlHelpURL: WdSmartTagControlTypeToString = "wdControlHelpURL"
        Case wdControlSeparator: WdSmartTagControlTypeToString = "wdControlSeparator"
        Case wdControlButton: WdSmartTagControlTypeToString = "wdControlButton"
        Case wdControlLabel: WdSmartTagControlTypeToString = "wdControlLabel"
        Case wdControlImage: WdSmartTagControlTypeToString = "wdControlImage"
        Case wdControlCheckbox: WdSmartTagControlTypeToString = "wdControlCheckbox"
        Case wdControlTextbox: WdSmartTagControlTypeToString = "wdControlTextbox"
        Case wdControlListbox: WdSmartTagControlTypeToString = "wdControlListbox"
        Case wdControlCombo: WdSmartTagControlTypeToString = "wdControlCombo"
        Case wdControlActiveX: WdSmartTagControlTypeToString = "wdControlActiveX"
        Case wdControlDocumentFragment: WdSmartTagControlTypeToString = "wdControlDocumentFragment"
        Case wdControlDocumentFragmentURL: WdSmartTagControlTypeToString = "wdControlDocumentFragmentURL"
        Case wdControlRadioGroup: WdSmartTagControlTypeToString = "wdControlRadioGroup"
    End Select
End Function
