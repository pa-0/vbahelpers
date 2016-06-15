Attribute VB_Name = "wOlAccountType"
Function OlAccountTypeFromString(value As String) As OlAccountType
    If IsNumeric(value) Then
        OlAccountTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olExchange": OlAccountTypeFromString = olExchange
        Case "olImap": OlAccountTypeFromString = olImap
        Case "olPop3": OlAccountTypeFromString = olPop3
        Case "olHttp": OlAccountTypeFromString = olHttp
        Case "olOtherAccount": OlAccountTypeFromString = olOtherAccount
    End Select
End Function

Function OlAccountTypeToString(value As OlAccountType) As String
    Select Case value
        Case olExchange: OlAccountTypeToString = "olExchange"
        Case olImap: OlAccountTypeToString = "olImap"
        Case olPop3: OlAccountTypeToString = "olPop3"
        Case olHttp: OlAccountTypeToString = "olHttp"
        Case olOtherAccount: OlAccountTypeToString = "olOtherAccount"
    End Select
End Function
