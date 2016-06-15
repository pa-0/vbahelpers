Attribute VB_Name = "wPbColorMode"
Function PbColorModeFromString(value As String) As PbColorMode
    If IsNumeric(value) Then
        PbColorModeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbColorModeDesktop": PbColorModeFromString = pbColorModeDesktop
        Case "pbColorModeProcess": PbColorModeFromString = pbColorModeProcess
        Case "pbColorModeSpot": PbColorModeFromString = pbColorModeSpot
        Case "pbColorModeBW": PbColorModeFromString = pbColorModeBW
        Case "pbColorModeSpotAndProcess": PbColorModeFromString = pbColorModeSpotAndProcess
    End Select
End Function

Function PbColorModeToString(value As PbColorMode) As String
    Select Case value
        Case pbColorModeDesktop: PbColorModeToString = "pbColorModeDesktop"
        Case pbColorModeProcess: PbColorModeToString = "pbColorModeProcess"
        Case pbColorModeSpot: PbColorModeToString = "pbColorModeSpot"
        Case pbColorModeBW: PbColorModeToString = "pbColorModeBW"
        Case pbColorModeSpotAndProcess: PbColorModeToString = "pbColorModeSpotAndProcess"
    End Select
End Function
