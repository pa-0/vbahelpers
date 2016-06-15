Attribute VB_Name = "wPpSlideShowState"
Function PpSlideShowStateFromString(value As String) As PpSlideShowState
    If IsNumeric(value) Then
        PpSlideShowStateFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "ppSlideShowRunning": PpSlideShowStateFromString = ppSlideShowRunning
        Case "ppSlideShowPaused": PpSlideShowStateFromString = ppSlideShowPaused
        Case "ppSlideShowBlackScreen": PpSlideShowStateFromString = ppSlideShowBlackScreen
        Case "ppSlideShowWhiteScreen": PpSlideShowStateFromString = ppSlideShowWhiteScreen
        Case "ppSlideShowDone": PpSlideShowStateFromString = ppSlideShowDone
    End Select
End Function

Function PpSlideShowStateToString(value As PpSlideShowState) As String
    Select Case value
        Case ppSlideShowRunning: PpSlideShowStateToString = "ppSlideShowRunning"
        Case ppSlideShowPaused: PpSlideShowStateToString = "ppSlideShowPaused"
        Case ppSlideShowBlackScreen: PpSlideShowStateToString = "ppSlideShowBlackScreen"
        Case ppSlideShowWhiteScreen: PpSlideShowStateToString = "ppSlideShowWhiteScreen"
        Case ppSlideShowDone: PpSlideShowStateToString = "ppSlideShowDone"
    End Select
End Function
