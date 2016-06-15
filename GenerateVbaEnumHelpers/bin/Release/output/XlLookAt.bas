Attribute VB_Name = "wXlLookAt"
Function XlLookAtFromString(value As String) As XlLookAt
    If IsNumeric(value) Then
        XlLookAtFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlWhole": XlLookAtFromString = xlWhole
        Case "xlPart": XlLookAtFromString = xlPart
    End Select
End Function

Function XlLookAtToString(value As XlLookAt) As String
    Select Case value
        Case xlWhole: XlLookAtToString = "xlWhole"
        Case xlPart: XlLookAtToString = "xlPart"
    End Select
End Function
