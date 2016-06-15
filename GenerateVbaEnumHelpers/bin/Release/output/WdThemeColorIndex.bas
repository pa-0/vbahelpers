Attribute VB_Name = "wWdThemeColorIndex"
Function WdThemeColorIndexFromString(value As String) As WdThemeColorIndex
    If IsNumeric(value) Then
        WdThemeColorIndexFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdThemeColorMainDark1": WdThemeColorIndexFromString = wdThemeColorMainDark1
        Case "wdThemeColorMainLight1": WdThemeColorIndexFromString = wdThemeColorMainLight1
        Case "wdThemeColorMainDark2": WdThemeColorIndexFromString = wdThemeColorMainDark2
        Case "wdThemeColorMainLight2": WdThemeColorIndexFromString = wdThemeColorMainLight2
        Case "wdThemeColorAccent1": WdThemeColorIndexFromString = wdThemeColorAccent1
        Case "wdThemeColorAccent2": WdThemeColorIndexFromString = wdThemeColorAccent2
        Case "wdThemeColorAccent3": WdThemeColorIndexFromString = wdThemeColorAccent3
        Case "wdThemeColorAccent4": WdThemeColorIndexFromString = wdThemeColorAccent4
        Case "wdThemeColorAccent5": WdThemeColorIndexFromString = wdThemeColorAccent5
        Case "wdThemeColorAccent6": WdThemeColorIndexFromString = wdThemeColorAccent6
        Case "wdThemeColorHyperlink": WdThemeColorIndexFromString = wdThemeColorHyperlink
        Case "wdThemeColorHyperlinkFollowed": WdThemeColorIndexFromString = wdThemeColorHyperlinkFollowed
        Case "wdThemeColorBackground1": WdThemeColorIndexFromString = wdThemeColorBackground1
        Case "wdThemeColorText1": WdThemeColorIndexFromString = wdThemeColorText1
        Case "wdThemeColorBackground2": WdThemeColorIndexFromString = wdThemeColorBackground2
        Case "wdThemeColorText2": WdThemeColorIndexFromString = wdThemeColorText2
        Case "wdNotThemeColor": WdThemeColorIndexFromString = wdNotThemeColor
    End Select
End Function

Function WdThemeColorIndexToString(value As WdThemeColorIndex) As String
    Select Case value
        Case wdThemeColorMainDark1: WdThemeColorIndexToString = "wdThemeColorMainDark1"
        Case wdThemeColorMainLight1: WdThemeColorIndexToString = "wdThemeColorMainLight1"
        Case wdThemeColorMainDark2: WdThemeColorIndexToString = "wdThemeColorMainDark2"
        Case wdThemeColorMainLight2: WdThemeColorIndexToString = "wdThemeColorMainLight2"
        Case wdThemeColorAccent1: WdThemeColorIndexToString = "wdThemeColorAccent1"
        Case wdThemeColorAccent2: WdThemeColorIndexToString = "wdThemeColorAccent2"
        Case wdThemeColorAccent3: WdThemeColorIndexToString = "wdThemeColorAccent3"
        Case wdThemeColorAccent4: WdThemeColorIndexToString = "wdThemeColorAccent4"
        Case wdThemeColorAccent5: WdThemeColorIndexToString = "wdThemeColorAccent5"
        Case wdThemeColorAccent6: WdThemeColorIndexToString = "wdThemeColorAccent6"
        Case wdThemeColorHyperlink: WdThemeColorIndexToString = "wdThemeColorHyperlink"
        Case wdThemeColorHyperlinkFollowed: WdThemeColorIndexToString = "wdThemeColorHyperlinkFollowed"
        Case wdThemeColorBackground1: WdThemeColorIndexToString = "wdThemeColorBackground1"
        Case wdThemeColorText1: WdThemeColorIndexToString = "wdThemeColorText1"
        Case wdThemeColorBackground2: WdThemeColorIndexToString = "wdThemeColorBackground2"
        Case wdThemeColorText2: WdThemeColorIndexToString = "wdThemeColorText2"
        Case wdNotThemeColor: WdThemeColorIndexToString = "wdNotThemeColor"
    End Select
End Function
