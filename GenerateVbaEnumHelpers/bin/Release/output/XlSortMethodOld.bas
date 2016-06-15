Attribute VB_Name = "wXlSortMethodOld"
Function XlSortMethodOldFromString(value As String) As XlSortMethodOld
    If IsNumeric(value) Then
        XlSortMethodOldFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlSyllabary": XlSortMethodOldFromString = xlSyllabary
        Case "xlCodePage": XlSortMethodOldFromString = xlCodePage
    End Select
End Function

Function XlSortMethodOldToString(value As XlSortMethodOld) As String
    Select Case value
        Case xlSyllabary: XlSortMethodOldToString = "xlSyllabary"
        Case xlCodePage: XlSortMethodOldToString = "xlCodePage"
    End Select
End Function
