Attribute VB_Name = "wOlFormatEnumeration"
Function OlFormatEnumerationFromString(value As String) As OlFormatEnumeration
    If IsNumeric(value) Then
        OlFormatEnumerationFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olFormatEnumBitmap": OlFormatEnumerationFromString = olFormatEnumBitmap
        Case "olFormatEnumText": OlFormatEnumerationFromString = olFormatEnumText
    End Select
End Function

Function OlFormatEnumerationToString(value As OlFormatEnumeration) As String
    Select Case value
        Case olFormatEnumBitmap: OlFormatEnumerationToString = "olFormatEnumBitmap"
        Case olFormatEnumText: OlFormatEnumerationToString = "olFormatEnumText"
    End Select
End Function
