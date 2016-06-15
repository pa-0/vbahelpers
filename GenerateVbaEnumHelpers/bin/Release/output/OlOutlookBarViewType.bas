Attribute VB_Name = "wOlOutlookBarViewType"
Function OlOutlookBarViewTypeFromString(value As String) As OlOutlookBarViewType
    If IsNumeric(value) Then
        OlOutlookBarViewTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olLargeIcon": OlOutlookBarViewTypeFromString = olLargeIcon
        Case "olSmallIcon": OlOutlookBarViewTypeFromString = olSmallIcon
    End Select
End Function

Function OlOutlookBarViewTypeToString(value As OlOutlookBarViewType) As String
    Select Case value
        Case olLargeIcon: OlOutlookBarViewTypeToString = "olLargeIcon"
        Case olSmallIcon: OlOutlookBarViewTypeToString = "olSmallIcon"
    End Select
End Function
