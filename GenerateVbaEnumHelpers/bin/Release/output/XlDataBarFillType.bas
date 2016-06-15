Attribute VB_Name = "wXlDataBarFillType"
Function XlDataBarFillTypeFromString(value As String) As XlDataBarFillType
    If IsNumeric(value) Then
        XlDataBarFillTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlDataBarFillSolid": XlDataBarFillTypeFromString = xlDataBarFillSolid
        Case "xlDataBarFillGradient": XlDataBarFillTypeFromString = xlDataBarFillGradient
    End Select
End Function

Function XlDataBarFillTypeToString(value As XlDataBarFillType) As String
    Select Case value
        Case xlDataBarFillSolid: XlDataBarFillTypeToString = "xlDataBarFillSolid"
        Case xlDataBarFillGradient: XlDataBarFillTypeToString = "xlDataBarFillGradient"
    End Select
End Function
