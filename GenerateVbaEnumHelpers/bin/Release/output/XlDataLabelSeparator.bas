Attribute VB_Name = "wXlDataLabelSeparator"
Function XlDataLabelSeparatorFromString(value As String) As XlDataLabelSeparator
    If IsNumeric(value) Then
        XlDataLabelSeparatorFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlDataLabelSeparatorDefault": XlDataLabelSeparatorFromString = xlDataLabelSeparatorDefault
    End Select
End Function

Function XlDataLabelSeparatorToString(value As XlDataLabelSeparator) As String
    Select Case value
        Case xlDataLabelSeparatorDefault: XlDataLabelSeparatorToString = "xlDataLabelSeparatorDefault"
    End Select
End Function
