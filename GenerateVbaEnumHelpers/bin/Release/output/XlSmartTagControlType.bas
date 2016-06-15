Attribute VB_Name = "wXlSmartTagControlType"
Function XlSmartTagControlTypeFromString(value As String) As XlSmartTagControlType
    If IsNumeric(value) Then
        XlSmartTagControlTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlSmartTagControlSmartTag": XlSmartTagControlTypeFromString = xlSmartTagControlSmartTag
        Case "xlSmartTagControlLink": XlSmartTagControlTypeFromString = xlSmartTagControlLink
        Case "xlSmartTagControlHelp": XlSmartTagControlTypeFromString = xlSmartTagControlHelp
        Case "xlSmartTagControlHelpURL": XlSmartTagControlTypeFromString = xlSmartTagControlHelpURL
        Case "xlSmartTagControlSeparator": XlSmartTagControlTypeFromString = xlSmartTagControlSeparator
        Case "xlSmartTagControlButton": XlSmartTagControlTypeFromString = xlSmartTagControlButton
        Case "xlSmartTagControlLabel": XlSmartTagControlTypeFromString = xlSmartTagControlLabel
        Case "xlSmartTagControlImage": XlSmartTagControlTypeFromString = xlSmartTagControlImage
        Case "xlSmartTagControlCheckbox": XlSmartTagControlTypeFromString = xlSmartTagControlCheckbox
        Case "xlSmartTagControlTextbox": XlSmartTagControlTypeFromString = xlSmartTagControlTextbox
        Case "xlSmartTagControlListbox": XlSmartTagControlTypeFromString = xlSmartTagControlListbox
        Case "xlSmartTagControlCombo": XlSmartTagControlTypeFromString = xlSmartTagControlCombo
        Case "xlSmartTagControlActiveX": XlSmartTagControlTypeFromString = xlSmartTagControlActiveX
        Case "xlSmartTagControlRadioGroup": XlSmartTagControlTypeFromString = xlSmartTagControlRadioGroup
    End Select
End Function

Function XlSmartTagControlTypeToString(value As XlSmartTagControlType) As String
    Select Case value
        Case xlSmartTagControlSmartTag: XlSmartTagControlTypeToString = "xlSmartTagControlSmartTag"
        Case xlSmartTagControlLink: XlSmartTagControlTypeToString = "xlSmartTagControlLink"
        Case xlSmartTagControlHelp: XlSmartTagControlTypeToString = "xlSmartTagControlHelp"
        Case xlSmartTagControlHelpURL: XlSmartTagControlTypeToString = "xlSmartTagControlHelpURL"
        Case xlSmartTagControlSeparator: XlSmartTagControlTypeToString = "xlSmartTagControlSeparator"
        Case xlSmartTagControlButton: XlSmartTagControlTypeToString = "xlSmartTagControlButton"
        Case xlSmartTagControlLabel: XlSmartTagControlTypeToString = "xlSmartTagControlLabel"
        Case xlSmartTagControlImage: XlSmartTagControlTypeToString = "xlSmartTagControlImage"
        Case xlSmartTagControlCheckbox: XlSmartTagControlTypeToString = "xlSmartTagControlCheckbox"
        Case xlSmartTagControlTextbox: XlSmartTagControlTypeToString = "xlSmartTagControlTextbox"
        Case xlSmartTagControlListbox: XlSmartTagControlTypeToString = "xlSmartTagControlListbox"
        Case xlSmartTagControlCombo: XlSmartTagControlTypeToString = "xlSmartTagControlCombo"
        Case xlSmartTagControlActiveX: XlSmartTagControlTypeToString = "xlSmartTagControlActiveX"
        Case xlSmartTagControlRadioGroup: XlSmartTagControlTypeToString = "xlSmartTagControlRadioGroup"
    End Select
End Function
