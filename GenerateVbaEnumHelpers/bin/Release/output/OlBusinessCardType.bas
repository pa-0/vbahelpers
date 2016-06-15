Attribute VB_Name = "wOlBusinessCardType"
Function OlBusinessCardTypeFromString(value As String) As OlBusinessCardType
    If IsNumeric(value) Then
        OlBusinessCardTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olBusinessCardTypeOutlook": OlBusinessCardTypeFromString = olBusinessCardTypeOutlook
        Case "olBusinessCardTypeInterConnect": OlBusinessCardTypeFromString = olBusinessCardTypeInterConnect
    End Select
End Function

Function OlBusinessCardTypeToString(value As OlBusinessCardType) As String
    Select Case value
        Case olBusinessCardTypeOutlook: OlBusinessCardTypeToString = "olBusinessCardTypeOutlook"
        Case olBusinessCardTypeInterConnect: OlBusinessCardTypeToString = "olBusinessCardTypeInterConnect"
    End Select
End Function
