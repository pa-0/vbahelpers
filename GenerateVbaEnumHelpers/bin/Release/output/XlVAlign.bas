Attribute VB_Name = "wXlVAlign"
Function XlVAlignFromString(value As String) As XlVAlign
    If IsNumeric(value) Then
        XlVAlignFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlVAlignTop": XlVAlignFromString = xlVAlignTop
        Case "xlVAlignJustify": XlVAlignFromString = xlVAlignJustify
        Case "xlVAlignDistributed": XlVAlignFromString = xlVAlignDistributed
        Case "xlVAlignCenter": XlVAlignFromString = xlVAlignCenter
        Case "xlVAlignBottom": XlVAlignFromString = xlVAlignBottom
    End Select
End Function

Function XlVAlignToString(value As XlVAlign) As String
    Select Case value
        Case xlVAlignTop: XlVAlignToString = "xlVAlignTop"
        Case xlVAlignJustify: XlVAlignToString = "xlVAlignJustify"
        Case xlVAlignDistributed: XlVAlignToString = "xlVAlignDistributed"
        Case xlVAlignCenter: XlVAlignToString = "xlVAlignCenter"
        Case xlVAlignBottom: XlVAlignToString = "xlVAlignBottom"
    End Select
End Function
