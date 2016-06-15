Attribute VB_Name = "wXlDataBarBorderType"
Function XlDataBarBorderTypeFromString(value As String) As XlDataBarBorderType
    If IsNumeric(value) Then
        XlDataBarBorderTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlDataBarBorderNone": XlDataBarBorderTypeFromString = xlDataBarBorderNone
        Case "xlDataBarBorderSolid": XlDataBarBorderTypeFromString = xlDataBarBorderSolid
    End Select
End Function

Function XlDataBarBorderTypeToString(value As XlDataBarBorderType) As String
    Select Case value
        Case xlDataBarBorderNone: XlDataBarBorderTypeToString = "xlDataBarBorderNone"
        Case xlDataBarBorderSolid: XlDataBarBorderTypeToString = "xlDataBarBorderSolid"
    End Select
End Function
