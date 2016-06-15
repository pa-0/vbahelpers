Attribute VB_Name = "wOlStoreType"
Function OlStoreTypeFromString(value As String) As OlStoreType
    If IsNumeric(value) Then
        OlStoreTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olStoreDefault": OlStoreTypeFromString = olStoreDefault
        Case "olStoreUnicode": OlStoreTypeFromString = olStoreUnicode
        Case "olStoreANSI": OlStoreTypeFromString = olStoreANSI
    End Select
End Function

Function OlStoreTypeToString(value As OlStoreType) As String
    Select Case value
        Case olStoreDefault: OlStoreTypeToString = "olStoreDefault"
        Case olStoreUnicode: OlStoreTypeToString = "olStoreUnicode"
        Case olStoreANSI: OlStoreTypeToString = "olStoreANSI"
    End Select
End Function
