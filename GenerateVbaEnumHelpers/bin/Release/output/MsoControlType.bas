Attribute VB_Name = "wMsoControlType"
Function MsoControlTypeFromString(value As String) As MsoControlType
    If IsNumeric(value) Then
        MsoControlTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoControlCustom": MsoControlTypeFromString = msoControlCustom
        Case "msoControlButton": MsoControlTypeFromString = msoControlButton
        Case "msoControlEdit": MsoControlTypeFromString = msoControlEdit
        Case "msoControlDropdown": MsoControlTypeFromString = msoControlDropdown
        Case "msoControlComboBox": MsoControlTypeFromString = msoControlComboBox
        Case "msoControlButtonDropdown": MsoControlTypeFromString = msoControlButtonDropdown
        Case "msoControlSplitDropdown": MsoControlTypeFromString = msoControlSplitDropdown
        Case "msoControlOCXDropdown": MsoControlTypeFromString = msoControlOCXDropdown
        Case "msoControlGenericDropdown": MsoControlTypeFromString = msoControlGenericDropdown
        Case "msoControlGraphicDropdown": MsoControlTypeFromString = msoControlGraphicDropdown
        Case "msoControlPopup": MsoControlTypeFromString = msoControlPopup
        Case "msoControlGraphicPopup": MsoControlTypeFromString = msoControlGraphicPopup
        Case "msoControlButtonPopup": MsoControlTypeFromString = msoControlButtonPopup
        Case "msoControlSplitButtonPopup": MsoControlTypeFromString = msoControlSplitButtonPopup
        Case "msoControlSplitButtonMRUPopup": MsoControlTypeFromString = msoControlSplitButtonMRUPopup
        Case "msoControlLabel": MsoControlTypeFromString = msoControlLabel
        Case "msoControlExpandingGrid": MsoControlTypeFromString = msoControlExpandingGrid
        Case "msoControlSplitExpandingGrid": MsoControlTypeFromString = msoControlSplitExpandingGrid
        Case "msoControlGrid": MsoControlTypeFromString = msoControlGrid
        Case "msoControlGauge": MsoControlTypeFromString = msoControlGauge
        Case "msoControlGraphicCombo": MsoControlTypeFromString = msoControlGraphicCombo
        Case "msoControlPane": MsoControlTypeFromString = msoControlPane
        Case "msoControlActiveX": MsoControlTypeFromString = msoControlActiveX
        Case "msoControlSpinner": MsoControlTypeFromString = msoControlSpinner
        Case "msoControlLabelEx": MsoControlTypeFromString = msoControlLabelEx
        Case "msoControlWorkPane": MsoControlTypeFromString = msoControlWorkPane
        Case "msoControlAutoCompleteCombo": MsoControlTypeFromString = msoControlAutoCompleteCombo
    End Select
End Function

Function MsoControlTypeToString(value As MsoControlType) As String
    Select Case value
        Case msoControlCustom: MsoControlTypeToString = "msoControlCustom"
        Case msoControlButton: MsoControlTypeToString = "msoControlButton"
        Case msoControlEdit: MsoControlTypeToString = "msoControlEdit"
        Case msoControlDropdown: MsoControlTypeToString = "msoControlDropdown"
        Case msoControlComboBox: MsoControlTypeToString = "msoControlComboBox"
        Case msoControlButtonDropdown: MsoControlTypeToString = "msoControlButtonDropdown"
        Case msoControlSplitDropdown: MsoControlTypeToString = "msoControlSplitDropdown"
        Case msoControlOCXDropdown: MsoControlTypeToString = "msoControlOCXDropdown"
        Case msoControlGenericDropdown: MsoControlTypeToString = "msoControlGenericDropdown"
        Case msoControlGraphicDropdown: MsoControlTypeToString = "msoControlGraphicDropdown"
        Case msoControlPopup: MsoControlTypeToString = "msoControlPopup"
        Case msoControlGraphicPopup: MsoControlTypeToString = "msoControlGraphicPopup"
        Case msoControlButtonPopup: MsoControlTypeToString = "msoControlButtonPopup"
        Case msoControlSplitButtonPopup: MsoControlTypeToString = "msoControlSplitButtonPopup"
        Case msoControlSplitButtonMRUPopup: MsoControlTypeToString = "msoControlSplitButtonMRUPopup"
        Case msoControlLabel: MsoControlTypeToString = "msoControlLabel"
        Case msoControlExpandingGrid: MsoControlTypeToString = "msoControlExpandingGrid"
        Case msoControlSplitExpandingGrid: MsoControlTypeToString = "msoControlSplitExpandingGrid"
        Case msoControlGrid: MsoControlTypeToString = "msoControlGrid"
        Case msoControlGauge: MsoControlTypeToString = "msoControlGauge"
        Case msoControlGraphicCombo: MsoControlTypeToString = "msoControlGraphicCombo"
        Case msoControlPane: MsoControlTypeToString = "msoControlPane"
        Case msoControlActiveX: MsoControlTypeToString = "msoControlActiveX"
        Case msoControlSpinner: MsoControlTypeToString = "msoControlSpinner"
        Case msoControlLabelEx: MsoControlTypeToString = "msoControlLabelEx"
        Case msoControlWorkPane: MsoControlTypeToString = "msoControlWorkPane"
        Case msoControlAutoCompleteCombo: MsoControlTypeToString = "msoControlAutoCompleteCombo"
    End Select
End Function
