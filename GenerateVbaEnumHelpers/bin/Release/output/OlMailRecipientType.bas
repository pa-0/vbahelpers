Attribute VB_Name = "wOlMailRecipientType"
Function OlMailRecipientTypeFromString(value As String) As OlMailRecipientType
    If IsNumeric(value) Then
        OlMailRecipientTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olOriginator": OlMailRecipientTypeFromString = olOriginator
        Case "olTo": OlMailRecipientTypeFromString = olTo
        Case "olCC": OlMailRecipientTypeFromString = olCC
        Case "olBCC": OlMailRecipientTypeFromString = olBCC
    End Select
End Function

Function OlMailRecipientTypeToString(value As OlMailRecipientType) As String
    Select Case value
        Case olOriginator: OlMailRecipientTypeToString = "olOriginator"
        Case olTo: OlMailRecipientTypeToString = "olTo"
        Case olCC: OlMailRecipientTypeToString = "olCC"
        Case olBCC: OlMailRecipientTypeToString = "olBCC"
    End Select
End Function
