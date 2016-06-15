Attribute VB_Name = "wXlSortMethod"
Function XlSortMethodFromString(value As String) As XlSortMethod
    If IsNumeric(value) Then
        XlSortMethodFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlPinYin": XlSortMethodFromString = xlPinYin
        Case "xlStroke": XlSortMethodFromString = xlStroke
    End Select
End Function

Function XlSortMethodToString(value As XlSortMethod) As String
    Select Case value
        Case xlPinYin: XlSortMethodToString = "xlPinYin"
        Case xlStroke: XlSortMethodToString = "xlStroke"
    End Select
End Function
