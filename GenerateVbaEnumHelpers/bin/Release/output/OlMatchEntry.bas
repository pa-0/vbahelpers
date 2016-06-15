Attribute VB_Name = "wOlMatchEntry"
Function OlMatchEntryFromString(value As String) As OlMatchEntry
    If IsNumeric(value) Then
        OlMatchEntryFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olMatchEntryFirstLetter": OlMatchEntryFromString = olMatchEntryFirstLetter
        Case "olMatchEntryComplete": OlMatchEntryFromString = olMatchEntryComplete
        Case "olMatchEntryNone": OlMatchEntryFromString = olMatchEntryNone
    End Select
End Function

Function OlMatchEntryToString(value As OlMatchEntry) As String
    Select Case value
        Case olMatchEntryFirstLetter: OlMatchEntryToString = "olMatchEntryFirstLetter"
        Case olMatchEntryComplete: OlMatchEntryToString = "olMatchEntryComplete"
        Case olMatchEntryNone: OlMatchEntryToString = "olMatchEntryNone"
    End Select
End Function
