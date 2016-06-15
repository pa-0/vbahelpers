Attribute VB_Name = "wXlSizeRepresents"
Function XlSizeRepresentsFromString(value As String) As XlSizeRepresents
    If IsNumeric(value) Then
        XlSizeRepresentsFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlSizeIsArea": XlSizeRepresentsFromString = xlSizeIsArea
        Case "xlSizeIsWidth": XlSizeRepresentsFromString = xlSizeIsWidth
    End Select
End Function

Function XlSizeRepresentsToString(value As XlSizeRepresents) As String
    Select Case value
        Case xlSizeIsArea: XlSizeRepresentsToString = "xlSizeIsArea"
        Case xlSizeIsWidth: XlSizeRepresentsToString = "xlSizeIsWidth"
    End Select
End Function
