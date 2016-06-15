Attribute VB_Name = "wOlRecipientSelectors"
Function OlRecipientSelectorsFromString(value As String) As OlRecipientSelectors
    If IsNumeric(value) Then
        OlRecipientSelectorsFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olShowNone": OlRecipientSelectorsFromString = olShowNone
        Case "olShowTo": OlRecipientSelectorsFromString = olShowTo
        Case "olShowToCc": OlRecipientSelectorsFromString = olShowToCc
        Case "olShowToCcBcc": OlRecipientSelectorsFromString = olShowToCcBcc
    End Select
End Function

Function OlRecipientSelectorsToString(value As OlRecipientSelectors) As String
    Select Case value
        Case olShowNone: OlRecipientSelectorsToString = "olShowNone"
        Case olShowTo: OlRecipientSelectorsToString = "olShowTo"
        Case olShowToCc: OlRecipientSelectorsToString = "olShowToCc"
        Case olShowToCcBcc: OlRecipientSelectorsToString = "olShowToCcBcc"
    End Select
End Function
