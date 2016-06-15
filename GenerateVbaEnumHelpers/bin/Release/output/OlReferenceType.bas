Attribute VB_Name = "wOlReferenceType"
Function OlReferenceTypeFromString(value As String) As OlReferenceType
    If IsNumeric(value) Then
        OlReferenceTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olWeak": OlReferenceTypeFromString = olWeak
        Case "olStrong": OlReferenceTypeFromString = olStrong
    End Select
End Function

Function OlReferenceTypeToString(value As OlReferenceType) As String
    Select Case value
        Case olWeak: OlReferenceTypeToString = "olWeak"
        Case olStrong: OlReferenceTypeToString = "olStrong"
    End Select
End Function
