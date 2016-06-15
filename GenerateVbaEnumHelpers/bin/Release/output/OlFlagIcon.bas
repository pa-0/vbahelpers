Attribute VB_Name = "wOlFlagIcon"
Function OlFlagIconFromString(value As String) As OlFlagIcon
    If IsNumeric(value) Then
        OlFlagIconFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olNoFlagIcon": OlFlagIconFromString = olNoFlagIcon
        Case "olPurpleFlagIcon": OlFlagIconFromString = olPurpleFlagIcon
        Case "olOrangeFlagIcon": OlFlagIconFromString = olOrangeFlagIcon
        Case "olGreenFlagIcon": OlFlagIconFromString = olGreenFlagIcon
        Case "olYellowFlagIcon": OlFlagIconFromString = olYellowFlagIcon
        Case "olBlueFlagIcon": OlFlagIconFromString = olBlueFlagIcon
        Case "olRedFlagIcon": OlFlagIconFromString = olRedFlagIcon
    End Select
End Function

Function OlFlagIconToString(value As OlFlagIcon) As String
    Select Case value
        Case olNoFlagIcon: OlFlagIconToString = "olNoFlagIcon"
        Case olPurpleFlagIcon: OlFlagIconToString = "olPurpleFlagIcon"
        Case olOrangeFlagIcon: OlFlagIconToString = "olOrangeFlagIcon"
        Case olGreenFlagIcon: OlFlagIconToString = "olGreenFlagIcon"
        Case olYellowFlagIcon: OlFlagIconToString = "olYellowFlagIcon"
        Case olBlueFlagIcon: OlFlagIconToString = "olBlueFlagIcon"
        Case olRedFlagIcon: OlFlagIconToString = "olRedFlagIcon"
    End Select
End Function
