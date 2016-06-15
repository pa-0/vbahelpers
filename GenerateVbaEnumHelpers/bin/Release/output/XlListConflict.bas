Attribute VB_Name = "wXlListConflict"
Function XlListConflictFromString(value As String) As XlListConflict
    If IsNumeric(value) Then
        XlListConflictFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlListConflictDialog": XlListConflictFromString = xlListConflictDialog
        Case "xlListConflictRetryAllConflicts": XlListConflictFromString = xlListConflictRetryAllConflicts
        Case "xlListConflictDiscardAllConflicts": XlListConflictFromString = xlListConflictDiscardAllConflicts
        Case "xlListConflictError": XlListConflictFromString = xlListConflictError
    End Select
End Function

Function XlListConflictToString(value As XlListConflict) As String
    Select Case value
        Case xlListConflictDialog: XlListConflictToString = "xlListConflictDialog"
        Case xlListConflictRetryAllConflicts: XlListConflictToString = "xlListConflictRetryAllConflicts"
        Case xlListConflictDiscardAllConflicts: XlListConflictToString = "xlListConflictDiscardAllConflicts"
        Case xlListConflictError: XlListConflictToString = "xlListConflictError"
    End Select
End Function
