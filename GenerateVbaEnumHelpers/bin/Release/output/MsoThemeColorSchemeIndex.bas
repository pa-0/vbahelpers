Attribute VB_Name = "wMsoThemeColorSchemeIndex"
Function MsoThemeColorSchemeIndexFromString(value As String) As MsoThemeColorSchemeIndex
    If IsNumeric(value) Then
        MsoThemeColorSchemeIndexFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoThemeDark1": MsoThemeColorSchemeIndexFromString = msoThemeDark1
        Case "msoThemeLight1": MsoThemeColorSchemeIndexFromString = msoThemeLight1
        Case "msoThemeDark2": MsoThemeColorSchemeIndexFromString = msoThemeDark2
        Case "msoThemeLight2": MsoThemeColorSchemeIndexFromString = msoThemeLight2
        Case "msoThemeAccent1": MsoThemeColorSchemeIndexFromString = msoThemeAccent1
        Case "msoThemeAccent2": MsoThemeColorSchemeIndexFromString = msoThemeAccent2
        Case "msoThemeAccent3": MsoThemeColorSchemeIndexFromString = msoThemeAccent3
        Case "msoThemeAccent4": MsoThemeColorSchemeIndexFromString = msoThemeAccent4
        Case "msoThemeAccent5": MsoThemeColorSchemeIndexFromString = msoThemeAccent5
        Case "msoThemeAccent6": MsoThemeColorSchemeIndexFromString = msoThemeAccent6
        Case "msoThemeHyperlink": MsoThemeColorSchemeIndexFromString = msoThemeHyperlink
        Case "msoThemeFollowedHyperlink": MsoThemeColorSchemeIndexFromString = msoThemeFollowedHyperlink
    End Select
End Function

Function MsoThemeColorSchemeIndexToString(value As MsoThemeColorSchemeIndex) As String
    Select Case value
        Case msoThemeDark1: MsoThemeColorSchemeIndexToString = "msoThemeDark1"
        Case msoThemeLight1: MsoThemeColorSchemeIndexToString = "msoThemeLight1"
        Case msoThemeDark2: MsoThemeColorSchemeIndexToString = "msoThemeDark2"
        Case msoThemeLight2: MsoThemeColorSchemeIndexToString = "msoThemeLight2"
        Case msoThemeAccent1: MsoThemeColorSchemeIndexToString = "msoThemeAccent1"
        Case msoThemeAccent2: MsoThemeColorSchemeIndexToString = "msoThemeAccent2"
        Case msoThemeAccent3: MsoThemeColorSchemeIndexToString = "msoThemeAccent3"
        Case msoThemeAccent4: MsoThemeColorSchemeIndexToString = "msoThemeAccent4"
        Case msoThemeAccent5: MsoThemeColorSchemeIndexToString = "msoThemeAccent5"
        Case msoThemeAccent6: MsoThemeColorSchemeIndexToString = "msoThemeAccent6"
        Case msoThemeHyperlink: MsoThemeColorSchemeIndexToString = "msoThemeHyperlink"
        Case msoThemeFollowedHyperlink: MsoThemeColorSchemeIndexToString = "msoThemeFollowedHyperlink"
    End Select
End Function
