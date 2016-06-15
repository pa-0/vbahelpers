Attribute VB_Name = "wPpColorSchemeIndex"
Function PpColorSchemeIndexFromString(value As String) As PpColorSchemeIndex
    If IsNumeric(value) Then
        PpColorSchemeIndexFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "ppNotSchemeColor": PpColorSchemeIndexFromString = ppNotSchemeColor
        Case "ppBackground": PpColorSchemeIndexFromString = ppBackground
        Case "ppForeground": PpColorSchemeIndexFromString = ppForeground
        Case "ppShadow": PpColorSchemeIndexFromString = ppShadow
        Case "ppTitle": PpColorSchemeIndexFromString = ppTitle
        Case "ppFill": PpColorSchemeIndexFromString = ppFill
        Case "ppAccent1": PpColorSchemeIndexFromString = ppAccent1
        Case "ppAccent2": PpColorSchemeIndexFromString = ppAccent2
        Case "ppAccent3": PpColorSchemeIndexFromString = ppAccent3
        Case "ppSchemeColorMixed": PpColorSchemeIndexFromString = ppSchemeColorMixed
    End Select
End Function

Function PpColorSchemeIndexToString(value As PpColorSchemeIndex) As String
    Select Case value
        Case ppNotSchemeColor: PpColorSchemeIndexToString = "ppNotSchemeColor"
        Case ppBackground: PpColorSchemeIndexToString = "ppBackground"
        Case ppForeground: PpColorSchemeIndexToString = "ppForeground"
        Case ppShadow: PpColorSchemeIndexToString = "ppShadow"
        Case ppTitle: PpColorSchemeIndexToString = "ppTitle"
        Case ppFill: PpColorSchemeIndexToString = "ppFill"
        Case ppAccent1: PpColorSchemeIndexToString = "ppAccent1"
        Case ppAccent2: PpColorSchemeIndexToString = "ppAccent2"
        Case ppAccent3: PpColorSchemeIndexToString = "ppAccent3"
        Case ppSchemeColorMixed: PpColorSchemeIndexToString = "ppSchemeColorMixed"
    End Select
End Function
