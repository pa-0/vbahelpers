Attribute VB_Name = "wContentVerificationResults"
Function ContentVerificationResultsFromString(value As String) As ContentVerificationResults
    If IsNumeric(value) Then
        ContentVerificationResultsFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "contverresError": ContentVerificationResultsFromString = contverresError
        Case "contverresVerifying": ContentVerificationResultsFromString = contverresVerifying
        Case "contverresUnverified": ContentVerificationResultsFromString = contverresUnverified
        Case "contverresValid": ContentVerificationResultsFromString = contverresValid
        Case "contverresModified": ContentVerificationResultsFromString = contverresModified
    End Select
End Function

Function ContentVerificationResultsToString(value As ContentVerificationResults) As String
    Select Case value
        Case contverresError: ContentVerificationResultsToString = "contverresError"
        Case contverresVerifying: ContentVerificationResultsToString = "contverresVerifying"
        Case contverresUnverified: ContentVerificationResultsToString = "contverresUnverified"
        Case contverresValid: ContentVerificationResultsToString = "contverresValid"
        Case contverresModified: ContentVerificationResultsToString = "contverresModified"
    End Select
End Function
