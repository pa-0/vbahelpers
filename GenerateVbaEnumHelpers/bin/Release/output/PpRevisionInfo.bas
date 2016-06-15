Attribute VB_Name = "wPpRevisionInfo"
Function PpRevisionInfoFromString(value As String) As PpRevisionInfo
    If IsNumeric(value) Then
        PpRevisionInfoFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "ppRevisionInfoNone": PpRevisionInfoFromString = ppRevisionInfoNone
        Case "ppRevisionInfoBaseline": PpRevisionInfoFromString = ppRevisionInfoBaseline
        Case "ppRevisionInfoMerged": PpRevisionInfoFromString = ppRevisionInfoMerged
    End Select
End Function

Function PpRevisionInfoToString(value As PpRevisionInfo) As String
    Select Case value
        Case ppRevisionInfoNone: PpRevisionInfoToString = "ppRevisionInfoNone"
        Case ppRevisionInfoBaseline: PpRevisionInfoToString = "ppRevisionInfoBaseline"
        Case ppRevisionInfoMerged: PpRevisionInfoToString = "ppRevisionInfoMerged"
    End Select
End Function
