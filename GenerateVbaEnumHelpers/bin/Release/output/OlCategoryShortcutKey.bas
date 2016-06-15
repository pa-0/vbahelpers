Attribute VB_Name = "wOlCategoryShortcutKey"
Function OlCategoryShortcutKeyFromString(value As String) As OlCategoryShortcutKey
    If IsNumeric(value) Then
        OlCategoryShortcutKeyFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olCategoryShortcutKeyNone": OlCategoryShortcutKeyFromString = olCategoryShortcutKeyNone
        Case "olCategoryShortcutKeyCtrlF2": OlCategoryShortcutKeyFromString = olCategoryShortcutKeyCtrlF2
        Case "olCategoryShortcutKeyCtrlF3": OlCategoryShortcutKeyFromString = olCategoryShortcutKeyCtrlF3
        Case "olCategoryShortcutKeyCtrlF4": OlCategoryShortcutKeyFromString = olCategoryShortcutKeyCtrlF4
        Case "olCategoryShortcutKeyCtrlF5": OlCategoryShortcutKeyFromString = olCategoryShortcutKeyCtrlF5
        Case "olCategoryShortcutKeyCtrlF6": OlCategoryShortcutKeyFromString = olCategoryShortcutKeyCtrlF6
        Case "olCategoryShortcutKeyCtrlF7": OlCategoryShortcutKeyFromString = olCategoryShortcutKeyCtrlF7
        Case "olCategoryShortcutKeyCtrlF8": OlCategoryShortcutKeyFromString = olCategoryShortcutKeyCtrlF8
        Case "olCategoryShortcutKeyCtrlF9": OlCategoryShortcutKeyFromString = olCategoryShortcutKeyCtrlF9
        Case "olCategoryShortcutKeyCtrlF10": OlCategoryShortcutKeyFromString = olCategoryShortcutKeyCtrlF10
        Case "olCategoryShortcutKeyCtrlF11": OlCategoryShortcutKeyFromString = olCategoryShortcutKeyCtrlF11
        Case "olCategoryShortcutKeyCtrlF12": OlCategoryShortcutKeyFromString = olCategoryShortcutKeyCtrlF12
    End Select
End Function

Function OlCategoryShortcutKeyToString(value As OlCategoryShortcutKey) As String
    Select Case value
        Case olCategoryShortcutKeyNone: OlCategoryShortcutKeyToString = "olCategoryShortcutKeyNone"
        Case olCategoryShortcutKeyCtrlF2: OlCategoryShortcutKeyToString = "olCategoryShortcutKeyCtrlF2"
        Case olCategoryShortcutKeyCtrlF3: OlCategoryShortcutKeyToString = "olCategoryShortcutKeyCtrlF3"
        Case olCategoryShortcutKeyCtrlF4: OlCategoryShortcutKeyToString = "olCategoryShortcutKeyCtrlF4"
        Case olCategoryShortcutKeyCtrlF5: OlCategoryShortcutKeyToString = "olCategoryShortcutKeyCtrlF5"
        Case olCategoryShortcutKeyCtrlF6: OlCategoryShortcutKeyToString = "olCategoryShortcutKeyCtrlF6"
        Case olCategoryShortcutKeyCtrlF7: OlCategoryShortcutKeyToString = "olCategoryShortcutKeyCtrlF7"
        Case olCategoryShortcutKeyCtrlF8: OlCategoryShortcutKeyToString = "olCategoryShortcutKeyCtrlF8"
        Case olCategoryShortcutKeyCtrlF9: OlCategoryShortcutKeyToString = "olCategoryShortcutKeyCtrlF9"
        Case olCategoryShortcutKeyCtrlF10: OlCategoryShortcutKeyToString = "olCategoryShortcutKeyCtrlF10"
        Case olCategoryShortcutKeyCtrlF11: OlCategoryShortcutKeyToString = "olCategoryShortcutKeyCtrlF11"
        Case olCategoryShortcutKeyCtrlF12: OlCategoryShortcutKeyToString = "olCategoryShortcutKeyCtrlF12"
    End Select
End Function
