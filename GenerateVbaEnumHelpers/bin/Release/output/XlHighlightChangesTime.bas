Attribute VB_Name = "wXlHighlightChangesTime"
Function XlHighlightChangesTimeFromString(value As String) As XlHighlightChangesTime
    If IsNumeric(value) Then
        XlHighlightChangesTimeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlSinceMyLastSave": XlHighlightChangesTimeFromString = xlSinceMyLastSave
        Case "xlAllChanges": XlHighlightChangesTimeFromString = xlAllChanges
        Case "xlNotYetReviewed": XlHighlightChangesTimeFromString = xlNotYetReviewed
    End Select
End Function

Function XlHighlightChangesTimeToString(value As XlHighlightChangesTime) As String
    Select Case value
        Case xlSinceMyLastSave: XlHighlightChangesTimeToString = "xlSinceMyLastSave"
        Case xlAllChanges: XlHighlightChangesTimeToString = "xlAllChanges"
        Case xlNotYetReviewed: XlHighlightChangesTimeToString = "xlNotYetReviewed"
    End Select
End Function
