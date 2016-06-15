Attribute VB_Name = "wXlUnderlineStyle"
Function XlUnderlineStyleFromString(value As String) As XlUnderlineStyle
    If IsNumeric(value) Then
        XlUnderlineStyleFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlUnderlineStyleSingle": XlUnderlineStyleFromString = xlUnderlineStyleSingle
        Case "xlUnderlineStyleSingleAccounting": XlUnderlineStyleFromString = xlUnderlineStyleSingleAccounting
        Case "xlUnderlineStyleDoubleAccounting": XlUnderlineStyleFromString = xlUnderlineStyleDoubleAccounting
        Case "xlUnderlineStyleNone": XlUnderlineStyleFromString = xlUnderlineStyleNone
        Case "xlUnderlineStyleDouble": XlUnderlineStyleFromString = xlUnderlineStyleDouble
    End Select
End Function

Function XlUnderlineStyleToString(value As XlUnderlineStyle) As String
    Select Case value
        Case xlUnderlineStyleSingle: XlUnderlineStyleToString = "xlUnderlineStyleSingle"
        Case xlUnderlineStyleSingleAccounting: XlUnderlineStyleToString = "xlUnderlineStyleSingleAccounting"
        Case xlUnderlineStyleDoubleAccounting: XlUnderlineStyleToString = "xlUnderlineStyleDoubleAccounting"
        Case xlUnderlineStyleNone: XlUnderlineStyleToString = "xlUnderlineStyleNone"
        Case xlUnderlineStyleDouble: XlUnderlineStyleToString = "xlUnderlineStyleDouble"
    End Select
End Function
