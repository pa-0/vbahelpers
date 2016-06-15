Attribute VB_Name = "wPbWebControlType"
Function PbWebControlTypeFromString(value As String) As PbWebControlType
    If IsNumeric(value) Then
        PbWebControlTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbWebControlCheckBox": PbWebControlTypeFromString = pbWebControlCheckBox
        Case "pbWebControlCommandButton": PbWebControlTypeFromString = pbWebControlCommandButton
        Case "pbWebControlListBox": PbWebControlTypeFromString = pbWebControlListBox
        Case "pbWebControlMultiLineTextBox": PbWebControlTypeFromString = pbWebControlMultiLineTextBox
        Case "pbWebControlOptionButton": PbWebControlTypeFromString = pbWebControlOptionButton
        Case "pbWebControlSingleLineTextBox": PbWebControlTypeFromString = pbWebControlSingleLineTextBox
        Case "pbWebControlWebComponent": PbWebControlTypeFromString = pbWebControlWebComponent
        Case "pbWebControlHTMLFragment": PbWebControlTypeFromString = pbWebControlHTMLFragment
        Case "pbWebControlHotSpot": PbWebControlTypeFromString = pbWebControlHotSpot
    End Select
End Function

Function PbWebControlTypeToString(value As PbWebControlType) As String
    Select Case value
        Case pbWebControlCheckBox: PbWebControlTypeToString = "pbWebControlCheckBox"
        Case pbWebControlCommandButton: PbWebControlTypeToString = "pbWebControlCommandButton"
        Case pbWebControlListBox: PbWebControlTypeToString = "pbWebControlListBox"
        Case pbWebControlMultiLineTextBox: PbWebControlTypeToString = "pbWebControlMultiLineTextBox"
        Case pbWebControlOptionButton: PbWebControlTypeToString = "pbWebControlOptionButton"
        Case pbWebControlSingleLineTextBox: PbWebControlTypeToString = "pbWebControlSingleLineTextBox"
        Case pbWebControlWebComponent: PbWebControlTypeToString = "pbWebControlWebComponent"
        Case pbWebControlHTMLFragment: PbWebControlTypeToString = "pbWebControlHTMLFragment"
        Case pbWebControlHotSpot: PbWebControlTypeToString = "pbWebControlHotSpot"
    End Select
End Function
