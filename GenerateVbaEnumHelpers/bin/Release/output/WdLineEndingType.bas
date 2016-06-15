Attribute VB_Name = "wWdLineEndingType"
Function WdLineEndingTypeFromString(value As String) As WdLineEndingType
    If IsNumeric(value) Then
        WdLineEndingTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdCRLF": WdLineEndingTypeFromString = wdCRLF
        Case "wdCROnly": WdLineEndingTypeFromString = wdCROnly
        Case "wdLFOnly": WdLineEndingTypeFromString = wdLFOnly
        Case "wdLFCR": WdLineEndingTypeFromString = wdLFCR
        Case "wdLSPS": WdLineEndingTypeFromString = wdLSPS
    End Select
End Function

Function WdLineEndingTypeToString(value As WdLineEndingType) As String
    Select Case value
        Case wdCRLF: WdLineEndingTypeToString = "wdCRLF"
        Case wdCROnly: WdLineEndingTypeToString = "wdCROnly"
        Case wdLFOnly: WdLineEndingTypeToString = "wdLFOnly"
        Case wdLFCR: WdLineEndingTypeToString = "wdLFCR"
        Case wdLSPS: WdLineEndingTypeToString = "wdLSPS"
    End Select
End Function
