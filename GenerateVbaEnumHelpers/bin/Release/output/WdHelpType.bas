Attribute VB_Name = "wWdHelpType"
Function WdHelpTypeFromString(value As String) As WdHelpType
    If IsNumeric(value) Then
        WdHelpTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdHelp": WdHelpTypeFromString = wdHelp
        Case "wdHelpAbout": WdHelpTypeFromString = wdHelpAbout
        Case "wdHelpActiveWindow": WdHelpTypeFromString = wdHelpActiveWindow
        Case "wdHelpContents": WdHelpTypeFromString = wdHelpContents
        Case "wdHelpExamplesAndDemos": WdHelpTypeFromString = wdHelpExamplesAndDemos
        Case "wdHelpIndex": WdHelpTypeFromString = wdHelpIndex
        Case "wdHelpKeyboard": WdHelpTypeFromString = wdHelpKeyboard
        Case "wdHelpPSSHelp": WdHelpTypeFromString = wdHelpPSSHelp
        Case "wdHelpQuickPreview": WdHelpTypeFromString = wdHelpQuickPreview
        Case "wdHelpSearch": WdHelpTypeFromString = wdHelpSearch
        Case "wdHelpUsingHelp": WdHelpTypeFromString = wdHelpUsingHelp
        Case "wdHelpIchitaro": WdHelpTypeFromString = wdHelpIchitaro
        Case "wdHelpPE2": WdHelpTypeFromString = wdHelpPE2
        Case "wdHelpHWP": WdHelpTypeFromString = wdHelpHWP
    End Select
End Function

Function WdHelpTypeToString(value As WdHelpType) As String
    Select Case value
        Case wdHelp: WdHelpTypeToString = "wdHelp"
        Case wdHelpAbout: WdHelpTypeToString = "wdHelpAbout"
        Case wdHelpActiveWindow: WdHelpTypeToString = "wdHelpActiveWindow"
        Case wdHelpContents: WdHelpTypeToString = "wdHelpContents"
        Case wdHelpExamplesAndDemos: WdHelpTypeToString = "wdHelpExamplesAndDemos"
        Case wdHelpIndex: WdHelpTypeToString = "wdHelpIndex"
        Case wdHelpKeyboard: WdHelpTypeToString = "wdHelpKeyboard"
        Case wdHelpPSSHelp: WdHelpTypeToString = "wdHelpPSSHelp"
        Case wdHelpQuickPreview: WdHelpTypeToString = "wdHelpQuickPreview"
        Case wdHelpSearch: WdHelpTypeToString = "wdHelpSearch"
        Case wdHelpUsingHelp: WdHelpTypeToString = "wdHelpUsingHelp"
        Case wdHelpIchitaro: WdHelpTypeToString = "wdHelpIchitaro"
        Case wdHelpPE2: WdHelpTypeToString = "wdHelpPE2"
        Case wdHelpHWP: WdHelpTypeToString = "wdHelpHWP"
    End Select
End Function
