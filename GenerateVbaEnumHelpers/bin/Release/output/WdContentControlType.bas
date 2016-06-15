Attribute VB_Name = "wWdContentControlType"
Function WdContentControlTypeFromString(value As String) As WdContentControlType
    If IsNumeric(value) Then
        WdContentControlTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdContentControlRichText": WdContentControlTypeFromString = wdContentControlRichText
        Case "wdContentControlText": WdContentControlTypeFromString = wdContentControlText
        Case "wdContentControlPicture": WdContentControlTypeFromString = wdContentControlPicture
        Case "wdContentControlComboBox": WdContentControlTypeFromString = wdContentControlComboBox
        Case "wdContentControlDropdownList": WdContentControlTypeFromString = wdContentControlDropdownList
        Case "wdContentControlBuildingBlockGallery": WdContentControlTypeFromString = wdContentControlBuildingBlockGallery
        Case "wdContentControlDate": WdContentControlTypeFromString = wdContentControlDate
        Case "wdContentControlGroup": WdContentControlTypeFromString = wdContentControlGroup
        Case "wdContentControlCheckBox": WdContentControlTypeFromString = wdContentControlCheckBox
    End Select
End Function

Function WdContentControlTypeToString(value As WdContentControlType) As String
    Select Case value
        Case wdContentControlRichText: WdContentControlTypeToString = "wdContentControlRichText"
        Case wdContentControlText: WdContentControlTypeToString = "wdContentControlText"
        Case wdContentControlPicture: WdContentControlTypeToString = "wdContentControlPicture"
        Case wdContentControlComboBox: WdContentControlTypeToString = "wdContentControlComboBox"
        Case wdContentControlDropdownList: WdContentControlTypeToString = "wdContentControlDropdownList"
        Case wdContentControlBuildingBlockGallery: WdContentControlTypeToString = "wdContentControlBuildingBlockGallery"
        Case wdContentControlDate: WdContentControlTypeToString = "wdContentControlDate"
        Case wdContentControlGroup: WdContentControlTypeToString = "wdContentControlGroup"
        Case wdContentControlCheckBox: WdContentControlTypeToString = "wdContentControlCheckBox"
    End Select
End Function
