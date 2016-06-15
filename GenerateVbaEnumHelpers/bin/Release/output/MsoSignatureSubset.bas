Attribute VB_Name = "wMsoSignatureSubset"
Function MsoSignatureSubsetFromString(value As String) As MsoSignatureSubset
    If IsNumeric(value) Then
        MsoSignatureSubsetFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoSignatureSubsetSignaturesAllSigs": MsoSignatureSubsetFromString = msoSignatureSubsetSignaturesAllSigs
        Case "msoSignatureSubsetSignaturesNonVisible": MsoSignatureSubsetFromString = msoSignatureSubsetSignaturesNonVisible
        Case "msoSignatureSubsetSignatureLines": MsoSignatureSubsetFromString = msoSignatureSubsetSignatureLines
        Case "msoSignatureSubsetSignatureLinesSigned": MsoSignatureSubsetFromString = msoSignatureSubsetSignatureLinesSigned
        Case "msoSignatureSubsetSignatureLinesUnsigned": MsoSignatureSubsetFromString = msoSignatureSubsetSignatureLinesUnsigned
        Case "msoSignatureSubsetAll": MsoSignatureSubsetFromString = msoSignatureSubsetAll
    End Select
End Function

Function MsoSignatureSubsetToString(value As MsoSignatureSubset) As String
    Select Case value
        Case msoSignatureSubsetSignaturesAllSigs: MsoSignatureSubsetToString = "msoSignatureSubsetSignaturesAllSigs"
        Case msoSignatureSubsetSignaturesNonVisible: MsoSignatureSubsetToString = "msoSignatureSubsetSignaturesNonVisible"
        Case msoSignatureSubsetSignatureLines: MsoSignatureSubsetToString = "msoSignatureSubsetSignatureLines"
        Case msoSignatureSubsetSignatureLinesSigned: MsoSignatureSubsetToString = "msoSignatureSubsetSignatureLinesSigned"
        Case msoSignatureSubsetSignatureLinesUnsigned: MsoSignatureSubsetToString = "msoSignatureSubsetSignatureLinesUnsigned"
        Case msoSignatureSubsetAll: MsoSignatureSubsetToString = "msoSignatureSubsetAll"
    End Select
End Function
