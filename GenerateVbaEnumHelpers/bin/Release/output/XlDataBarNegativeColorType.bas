Attribute VB_Name = "wXlDataBarNegativeColorType"
Function XlDataBarNegativeColorTypeFromString(value As String) As XlDataBarNegativeColorType
    If IsNumeric(value) Then
        XlDataBarNegativeColorTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlDataBarColor": XlDataBarNegativeColorTypeFromString = xlDataBarColor
        Case "xlDataBarSameAsPositive": XlDataBarNegativeColorTypeFromString = xlDataBarSameAsPositive
    End Select
End Function

Function XlDataBarNegativeColorTypeToString(value As XlDataBarNegativeColorType) As String
    Select Case value
        Case xlDataBarColor: XlDataBarNegativeColorTypeToString = "xlDataBarColor"
        Case xlDataBarSameAsPositive: XlDataBarNegativeColorTypeToString = "xlDataBarSameAsPositive"
    End Select
End Function
