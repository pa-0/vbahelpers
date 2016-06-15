Attribute VB_Name = "wXlReferenceStyle"
Function XlReferenceStyleFromString(value As String) As XlReferenceStyle
    If IsNumeric(value) Then
        XlReferenceStyleFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlA1": XlReferenceStyleFromString = xlA1
        Case "xlR1C1": XlReferenceStyleFromString = xlR1C1
    End Select
End Function

Function XlReferenceStyleToString(value As XlReferenceStyle) As String
    Select Case value
        Case xlA1: XlReferenceStyleToString = "xlA1"
        Case xlR1C1: XlReferenceStyleToString = "xlR1C1"
    End Select
End Function
