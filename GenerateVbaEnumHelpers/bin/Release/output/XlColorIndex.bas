Attribute VB_Name = "wXlColorIndex"
Function XlColorIndexFromString(value As String) As XlColorIndex
    If IsNumeric(value) Then
        XlColorIndexFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlColorIndexNone": XlColorIndexFromString = xlColorIndexNone
        Case "xlColorIndexAutomatic": XlColorIndexFromString = xlColorIndexAutomatic
    End Select
End Function

Function XlColorIndexToString(value As XlColorIndex) As String
    Select Case value
        Case xlColorIndexNone: XlColorIndexToString = "xlColorIndexNone"
        Case xlColorIndexAutomatic: XlColorIndexToString = "xlColorIndexAutomatic"
    End Select
End Function
