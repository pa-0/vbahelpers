Attribute VB_Name = "wPpTransitionSpeed"
Function PpTransitionSpeedFromString(value As String) As PpTransitionSpeed
    If IsNumeric(value) Then
        PpTransitionSpeedFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "ppTransitionSpeedSlow": PpTransitionSpeedFromString = ppTransitionSpeedSlow
        Case "ppTransitionSpeedMedium": PpTransitionSpeedFromString = ppTransitionSpeedMedium
        Case "ppTransitionSpeedFast": PpTransitionSpeedFromString = ppTransitionSpeedFast
        Case "ppTransitionSpeedMixed": PpTransitionSpeedFromString = ppTransitionSpeedMixed
    End Select
End Function

Function PpTransitionSpeedToString(value As PpTransitionSpeed) As String
    Select Case value
        Case ppTransitionSpeedSlow: PpTransitionSpeedToString = "ppTransitionSpeedSlow"
        Case ppTransitionSpeedMedium: PpTransitionSpeedToString = "ppTransitionSpeedMedium"
        Case ppTransitionSpeedFast: PpTransitionSpeedToString = "ppTransitionSpeedFast"
        Case ppTransitionSpeedMixed: PpTransitionSpeedToString = "ppTransitionSpeedMixed"
    End Select
End Function
