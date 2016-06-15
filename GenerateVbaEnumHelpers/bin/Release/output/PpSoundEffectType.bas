Attribute VB_Name = "wPpSoundEffectType"
Function PpSoundEffectTypeFromString(value As String) As PpSoundEffectType
    If IsNumeric(value) Then
        PpSoundEffectTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "ppSoundNone": PpSoundEffectTypeFromString = ppSoundNone
        Case "ppSoundStopPrevious": PpSoundEffectTypeFromString = ppSoundStopPrevious
        Case "ppSoundFile": PpSoundEffectTypeFromString = ppSoundFile
        Case "ppSoundEffectsMixed": PpSoundEffectTypeFromString = ppSoundEffectsMixed
    End Select
End Function

Function PpSoundEffectTypeToString(value As PpSoundEffectType) As String
    Select Case value
        Case ppSoundNone: PpSoundEffectTypeToString = "ppSoundNone"
        Case ppSoundStopPrevious: PpSoundEffectTypeToString = "ppSoundStopPrevious"
        Case ppSoundFile: PpSoundEffectTypeToString = "ppSoundFile"
        Case ppSoundEffectsMixed: PpSoundEffectTypeToString = "ppSoundEffectsMixed"
    End Select
End Function
