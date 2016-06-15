Attribute VB_Name = "wXlFillWith"
Function XlFillWithFromString(value As String) As XlFillWith
    If IsNumeric(value) Then
        XlFillWithFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlFillWithContents": XlFillWithFromString = xlFillWithContents
        Case "xlFillWithFormats": XlFillWithFromString = xlFillWithFormats
        Case "xlFillWithAll": XlFillWithFromString = xlFillWithAll
    End Select
End Function

Function XlFillWithToString(value As XlFillWith) As String
    Select Case value
        Case xlFillWithContents: XlFillWithToString = "xlFillWithContents"
        Case xlFillWithFormats: XlFillWithToString = "xlFillWithFormats"
        Case xlFillWithAll: XlFillWithToString = "xlFillWithAll"
    End Select
End Function
