Attribute VB_Name = "wMsoThemeColorIndex"
Function MsoThemeColorIndexFromString(value As String) As MsoThemeColorIndex
    If IsNumeric(value) Then
        MsoThemeColorIndexFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoNotThemeColor": MsoThemeColorIndexFromString = msoNotThemeColor
        Case "msoThemeColorDark1": MsoThemeColorIndexFromString = msoThemeColorDark1
        Case "msoThemeColorLight1": MsoThemeColorIndexFromString = msoThemeColorLight1
        Case "msoThemeColorDark2": MsoThemeColorIndexFromString = msoThemeColorDark2
        Case "msoThemeColorLight2": MsoThemeColorIndexFromString = msoThemeColorLight2
        Case "msoThemeColorAccent1": MsoThemeColorIndexFromString = msoThemeColorAccent1
        Case "msoThemeColorAccent2": MsoThemeColorIndexFromString = msoThemeColorAccent2
        Case "msoThemeColorAccent3": MsoThemeColorIndexFromString = msoThemeColorAccent3
        Case "msoThemeColorAccent4": MsoThemeColorIndexFromString = msoThemeColorAccent4
        Case "msoThemeColorAccent5": MsoThemeColorIndexFromString = msoThemeColorAccent5
        Case "msoThemeColorAccent6": MsoThemeColorIndexFromString = msoThemeColorAccent6
        Case "msoThemeColorHyperlink": MsoThemeColorIndexFromString = msoThemeColorHyperlink
        Case "msoThemeColorFollowedHyperlink": MsoThemeColorIndexFromString = msoThemeColorFollowedHyperlink
        Case "msoThemeColorText1": MsoThemeColorIndexFromString = msoThemeColorText1
        Case "msoThemeColorBackground1": MsoThemeColorIndexFromString = msoThemeColorBackground1
        Case "msoThemeColorText2": MsoThemeColorIndexFromString = msoThemeColorText2
        Case "msoThemeColorBackground2": MsoThemeColorIndexFromString = msoThemeColorBackground2
        Case "msoThemeColorMixed": MsoThemeColorIndexFromString = msoThemeColorMixed
    End Select
End Function

Function MsoThemeColorIndexToString(value As MsoThemeColorIndex) As String
    Select Case value
        Case msoNotThemeColor: MsoThemeColorIndexToString = "msoNotThemeColor"
        Case msoThemeColorDark1: MsoThemeColorIndexToString = "msoThemeColorDark1"
        Case msoThemeColorLight1: MsoThemeColorIndexToString = "msoThemeColorLight1"
        Case msoThemeColorDark2: MsoThemeColorIndexToString = "msoThemeColorDark2"
        Case msoThemeColorLight2: MsoThemeColorIndexToString = "msoThemeColorLight2"
        Case msoThemeColorAccent1: MsoThemeColorIndexToString = "msoThemeColorAccent1"
        Case msoThemeColorAccent2: MsoThemeColorIndexToString = "msoThemeColorAccent2"
        Case msoThemeColorAccent3: MsoThemeColorIndexToString = "msoThemeColorAccent3"
        Case msoThemeColorAccent4: MsoThemeColorIndexToString = "msoThemeColorAccent4"
        Case msoThemeColorAccent5: MsoThemeColorIndexToString = "msoThemeColorAccent5"
        Case msoThemeColorAccent6: MsoThemeColorIndexToString = "msoThemeColorAccent6"
        Case msoThemeColorHyperlink: MsoThemeColorIndexToString = "msoThemeColorHyperlink"
        Case msoThemeColorFollowedHyperlink: MsoThemeColorIndexToString = "msoThemeColorFollowedHyperlink"
        Case msoThemeColorText1: MsoThemeColorIndexToString = "msoThemeColorText1"
        Case msoThemeColorBackground1: MsoThemeColorIndexToString = "msoThemeColorBackground1"
        Case msoThemeColorText2: MsoThemeColorIndexToString = "msoThemeColorText2"
        Case msoThemeColorBackground2: MsoThemeColorIndexToString = "msoThemeColorBackground2"
        Case msoThemeColorMixed: MsoThemeColorIndexToString = "msoThemeColorMixed"
    End Select
End Function
