Attribute VB_Name = "wSignatureType"
Function SignatureTypeFromString(value As String) As SignatureType
    If IsNumeric(value) Then
        SignatureTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "sigtypeUnknown": SignatureTypeFromString = sigtypeUnknown
        Case "sigtypeNonVisible": SignatureTypeFromString = sigtypeNonVisible
        Case "sigtypeSignatureLine": SignatureTypeFromString = sigtypeSignatureLine
        Case "sigtypeMax": SignatureTypeFromString = sigtypeMax
    End Select
End Function

Function SignatureTypeToString(value As SignatureType) As String
    Select Case value
        Case sigtypeUnknown: SignatureTypeToString = "sigtypeUnknown"
        Case sigtypeNonVisible: SignatureTypeToString = "sigtypeNonVisible"
        Case sigtypeSignatureLine: SignatureTypeToString = "sigtypeSignatureLine"
        Case sigtypeMax: SignatureTypeToString = "sigtypeMax"
    End Select
End Function
