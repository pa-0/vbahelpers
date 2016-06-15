Attribute VB_Name = "wXlLayoutFormType"
Function XlLayoutFormTypeFromString(value As String) As XlLayoutFormType
    If IsNumeric(value) Then
        XlLayoutFormTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlTabular": XlLayoutFormTypeFromString = xlTabular
        Case "xlOutline": XlLayoutFormTypeFromString = xlOutline
    End Select
End Function

Function XlLayoutFormTypeToString(value As XlLayoutFormType) As String
    Select Case value
        Case xlTabular: XlLayoutFormTypeToString = "xlTabular"
        Case xlOutline: XlLayoutFormTypeToString = "xlOutline"
    End Select
End Function
