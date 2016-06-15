Attribute VB_Name = "wPpFollowColors"
Function PpFollowColorsFromString(value As String) As PpFollowColors
    If IsNumeric(value) Then
        PpFollowColorsFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "ppFollowColorsNone": PpFollowColorsFromString = ppFollowColorsNone
        Case "ppFollowColorsScheme": PpFollowColorsFromString = ppFollowColorsScheme
        Case "ppFollowColorsTextAndBackground": PpFollowColorsFromString = ppFollowColorsTextAndBackground
        Case "ppFollowColorsMixed": PpFollowColorsFromString = ppFollowColorsMixed
    End Select
End Function

Function PpFollowColorsToString(value As PpFollowColors) As String
    Select Case value
        Case ppFollowColorsNone: PpFollowColorsToString = "ppFollowColorsNone"
        Case ppFollowColorsScheme: PpFollowColorsToString = "ppFollowColorsScheme"
        Case ppFollowColorsTextAndBackground: PpFollowColorsToString = "ppFollowColorsTextAndBackground"
        Case ppFollowColorsMixed: PpFollowColorsToString = "ppFollowColorsMixed"
    End Select
End Function
