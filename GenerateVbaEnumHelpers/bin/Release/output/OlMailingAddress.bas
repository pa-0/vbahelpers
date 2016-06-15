Attribute VB_Name = "wOlMailingAddress"
Function OlMailingAddressFromString(value As String) As OlMailingAddress
    If IsNumeric(value) Then
        OlMailingAddressFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olNone": OlMailingAddressFromString = olNone
        Case "olHome": OlMailingAddressFromString = olHome
        Case "olBusiness": OlMailingAddressFromString = olBusiness
        Case "olOther": OlMailingAddressFromString = olOther
    End Select
End Function

Function OlMailingAddressToString(value As OlMailingAddress) As String
    Select Case value
        Case olNone: OlMailingAddressToString = "olNone"
        Case olHome: OlMailingAddressToString = "olHome"
        Case olBusiness: OlMailingAddressToString = "olBusiness"
        Case olOther: OlMailingAddressToString = "olOther"
    End Select
End Function
