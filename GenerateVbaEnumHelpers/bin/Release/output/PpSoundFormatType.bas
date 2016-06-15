Attribute VB_Name = "wPpSoundFormatType"
Function PpSoundFormatTypeFromString(value As String) As PpSoundFormatType
    If IsNumeric(value) Then
        PpSoundFormatTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "ppSoundFormatNone": PpSoundFormatTypeFromString = ppSoundFormatNone
        Case "ppSoundFormatWAV": PpSoundFormatTypeFromString = ppSoundFormatWAV
        Case "ppSoundFormatMIDI": PpSoundFormatTypeFromString = ppSoundFormatMIDI
        Case "ppSoundFormatCDAudio": PpSoundFormatTypeFromString = ppSoundFormatCDAudio
        Case "ppSoundFormatMixed": PpSoundFormatTypeFromString = ppSoundFormatMixed
    End Select
End Function

Function PpSoundFormatTypeToString(value As PpSoundFormatType) As String
    Select Case value
        Case ppSoundFormatNone: PpSoundFormatTypeToString = "ppSoundFormatNone"
        Case ppSoundFormatWAV: PpSoundFormatTypeToString = "ppSoundFormatWAV"
        Case ppSoundFormatMIDI: PpSoundFormatTypeToString = "ppSoundFormatMIDI"
        Case ppSoundFormatCDAudio: PpSoundFormatTypeToString = "ppSoundFormatCDAudio"
        Case ppSoundFormatMixed: PpSoundFormatTypeToString = "ppSoundFormatMixed"
    End Select
End Function
