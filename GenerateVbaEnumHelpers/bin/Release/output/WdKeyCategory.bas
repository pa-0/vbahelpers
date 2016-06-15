Attribute VB_Name = "wWdKeyCategory"
Function WdKeyCategoryFromString(value As String) As WdKeyCategory
    If IsNumeric(value) Then
        WdKeyCategoryFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdKeyCategoryDisable": WdKeyCategoryFromString = wdKeyCategoryDisable
        Case "wdKeyCategoryCommand": WdKeyCategoryFromString = wdKeyCategoryCommand
        Case "wdKeyCategoryMacro": WdKeyCategoryFromString = wdKeyCategoryMacro
        Case "wdKeyCategoryFont": WdKeyCategoryFromString = wdKeyCategoryFont
        Case "wdKeyCategoryAutoText": WdKeyCategoryFromString = wdKeyCategoryAutoText
        Case "wdKeyCategoryStyle": WdKeyCategoryFromString = wdKeyCategoryStyle
        Case "wdKeyCategorySymbol": WdKeyCategoryFromString = wdKeyCategorySymbol
        Case "wdKeyCategoryPrefix": WdKeyCategoryFromString = wdKeyCategoryPrefix
        Case "wdKeyCategoryNil": WdKeyCategoryFromString = wdKeyCategoryNil
    End Select
End Function

Function WdKeyCategoryToString(value As WdKeyCategory) As String
    Select Case value
        Case wdKeyCategoryDisable: WdKeyCategoryToString = "wdKeyCategoryDisable"
        Case wdKeyCategoryCommand: WdKeyCategoryToString = "wdKeyCategoryCommand"
        Case wdKeyCategoryMacro: WdKeyCategoryToString = "wdKeyCategoryMacro"
        Case wdKeyCategoryFont: WdKeyCategoryToString = "wdKeyCategoryFont"
        Case wdKeyCategoryAutoText: WdKeyCategoryToString = "wdKeyCategoryAutoText"
        Case wdKeyCategoryStyle: WdKeyCategoryToString = "wdKeyCategoryStyle"
        Case wdKeyCategorySymbol: WdKeyCategoryToString = "wdKeyCategorySymbol"
        Case wdKeyCategoryPrefix: WdKeyCategoryToString = "wdKeyCategoryPrefix"
        Case wdKeyCategoryNil: WdKeyCategoryToString = "wdKeyCategoryNil"
    End Select
End Function
