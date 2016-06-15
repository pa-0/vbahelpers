Attribute VB_Name = "wPpPlayerState"
Function PpPlayerStateFromString(value As String) As PpPlayerState
    If IsNumeric(value) Then
        PpPlayerStateFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "ppPlaying": PpPlayerStateFromString = ppPlaying
        Case "ppPaused": PpPlayerStateFromString = ppPaused
        Case "ppStopped": PpPlayerStateFromString = ppStopped
        Case "ppNotReady": PpPlayerStateFromString = ppNotReady
    End Select
End Function

Function PpPlayerStateToString(value As PpPlayerState) As String
    Select Case value
        Case ppPlaying: PpPlayerStateToString = "ppPlaying"
        Case ppPaused: PpPlayerStateToString = "ppPaused"
        Case ppStopped: PpPlayerStateToString = "ppStopped"
        Case ppNotReady: PpPlayerStateToString = "ppNotReady"
    End Select
End Function
