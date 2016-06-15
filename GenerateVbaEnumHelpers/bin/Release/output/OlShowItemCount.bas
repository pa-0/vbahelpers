Attribute VB_Name = "wOlShowItemCount"
Function OlShowItemCountFromString(value As String) As OlShowItemCount
    If IsNumeric(value) Then
        OlShowItemCountFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olNoItemCount": OlShowItemCountFromString = olNoItemCount
        Case "olShowUnreadItemCount": OlShowItemCountFromString = olShowUnreadItemCount
        Case "olShowTotalItemCount": OlShowItemCountFromString = olShowTotalItemCount
    End Select
End Function

Function OlShowItemCountToString(value As OlShowItemCount) As String
    Select Case value
        Case olNoItemCount: OlShowItemCountToString = "olNoItemCount"
        Case olShowUnreadItemCount: OlShowItemCountToString = "olShowUnreadItemCount"
        Case olShowTotalItemCount: OlShowItemCountToString = "olShowTotalItemCount"
    End Select
End Function
