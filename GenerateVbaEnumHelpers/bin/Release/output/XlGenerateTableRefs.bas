Attribute VB_Name = "wXlGenerateTableRefs"
Function XlGenerateTableRefsFromString(value As String) As XlGenerateTableRefs
    If IsNumeric(value) Then
        XlGenerateTableRefsFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlGenerateTableRefA1": XlGenerateTableRefsFromString = xlGenerateTableRefA1
        Case "xlGenerateTableRefStruct": XlGenerateTableRefsFromString = xlGenerateTableRefStruct
    End Select
End Function

Function XlGenerateTableRefsToString(value As XlGenerateTableRefs) As String
    Select Case value
        Case xlGenerateTableRefA1: XlGenerateTableRefsToString = "xlGenerateTableRefA1"
        Case xlGenerateTableRefStruct: XlGenerateTableRefsToString = "xlGenerateTableRefStruct"
    End Select
End Function
