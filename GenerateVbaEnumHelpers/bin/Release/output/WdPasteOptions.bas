Attribute VB_Name = "wWdPasteOptions"
Function WdPasteOptionsFromString(value As String) As WdPasteOptions
    If IsNumeric(value) Then
        WdPasteOptionsFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdKeepSourceFormatting": WdPasteOptionsFromString = wdKeepSourceFormatting
        Case "wdMatchDestinationFormatting": WdPasteOptionsFromString = wdMatchDestinationFormatting
        Case "wdKeepTextOnly": WdPasteOptionsFromString = wdKeepTextOnly
        Case "wdUseDestinationStyles": WdPasteOptionsFromString = wdUseDestinationStyles
    End Select
End Function

Function WdPasteOptionsToString(value As WdPasteOptions) As String
    Select Case value
        Case wdKeepSourceFormatting: WdPasteOptionsToString = "wdKeepSourceFormatting"
        Case wdMatchDestinationFormatting: WdPasteOptionsToString = "wdMatchDestinationFormatting"
        Case wdKeepTextOnly: WdPasteOptionsToString = "wdKeepTextOnly"
        Case wdUseDestinationStyles: WdPasteOptionsToString = "wdUseDestinationStyles"
    End Select
End Function
