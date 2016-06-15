Attribute VB_Name = "wXlThemeColor"
Function XlThemeColorFromString(value As String) As XlThemeColor
    If IsNumeric(value) Then
        XlThemeColorFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlThemeColorDark1": XlThemeColorFromString = xlThemeColorDark1
        Case "xlThemeColorLight1": XlThemeColorFromString = xlThemeColorLight1
        Case "xlThemeColorDark2": XlThemeColorFromString = xlThemeColorDark2
        Case "xlThemeColorLight2": XlThemeColorFromString = xlThemeColorLight2
        Case "xlThemeColorAccent1": XlThemeColorFromString = xlThemeColorAccent1
        Case "xlThemeColorAccent2": XlThemeColorFromString = xlThemeColorAccent2
        Case "xlThemeColorAccent3": XlThemeColorFromString = xlThemeColorAccent3
        Case "xlThemeColorAccent4": XlThemeColorFromString = xlThemeColorAccent4
        Case "xlThemeColorAccent5": XlThemeColorFromString = xlThemeColorAccent5
        Case "xlThemeColorAccent6": XlThemeColorFromString = xlThemeColorAccent6
        Case "xlThemeColorHyperlink": XlThemeColorFromString = xlThemeColorHyperlink
        Case "xlThemeColorFollowedHyperlink": XlThemeColorFromString = xlThemeColorFollowedHyperlink
    End Select
End Function

Function XlThemeColorToString(value As XlThemeColor) As String
    Select Case value
        Case xlThemeColorDark1: XlThemeColorToString = "xlThemeColorDark1"
        Case xlThemeColorLight1: XlThemeColorToString = "xlThemeColorLight1"
        Case xlThemeColorDark2: XlThemeColorToString = "xlThemeColorDark2"
        Case xlThemeColorLight2: XlThemeColorToString = "xlThemeColorLight2"
        Case xlThemeColorAccent1: XlThemeColorToString = "xlThemeColorAccent1"
        Case xlThemeColorAccent2: XlThemeColorToString = "xlThemeColorAccent2"
        Case xlThemeColorAccent3: XlThemeColorToString = "xlThemeColorAccent3"
        Case xlThemeColorAccent4: XlThemeColorToString = "xlThemeColorAccent4"
        Case xlThemeColorAccent5: XlThemeColorToString = "xlThemeColorAccent5"
        Case xlThemeColorAccent6: XlThemeColorToString = "xlThemeColorAccent6"
        Case xlThemeColorHyperlink: XlThemeColorToString = "xlThemeColorHyperlink"
        Case xlThemeColorFollowedHyperlink: XlThemeColorToString = "xlThemeColorFollowedHyperlink"
    End Select
End Function
