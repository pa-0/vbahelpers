Attribute VB_Name = "wOlBodyFormat"
Function OlBodyFormatFromString(value As String) As OlBodyFormat
    If IsNumeric(value) Then
        OlBodyFormatFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olFormatUnspecified": OlBodyFormatFromString = olFormatUnspecified
        Case "olFormatPlain": OlBodyFormatFromString = olFormatPlain
        Case "olFormatHTML": OlBodyFormatFromString = olFormatHTML
        Case "olFormatRichText": OlBodyFormatFromString = olFormatRichText
    End Select
End Function

Function OlBodyFormatToString(value As OlBodyFormat) As String
    Select Case value
        Case olFormatUnspecified: OlBodyFormatToString = "olFormatUnspecified"
        Case olFormatPlain: OlBodyFormatToString = "olFormatPlain"
        Case olFormatHTML: OlBodyFormatToString = "olFormatHTML"
        Case olFormatRichText: OlBodyFormatToString = "olFormatRichText"
    End Select
End Function
