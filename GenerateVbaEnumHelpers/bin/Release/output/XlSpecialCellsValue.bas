Attribute VB_Name = "wXlSpecialCellsValue"
Function XlSpecialCellsValueFromString(value As String) As XlSpecialCellsValue
    If IsNumeric(value) Then
        XlSpecialCellsValueFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlNumbers": XlSpecialCellsValueFromString = xlNumbers
        Case "xlTextValues": XlSpecialCellsValueFromString = xlTextValues
        Case "xlLogical": XlSpecialCellsValueFromString = xlLogical
        Case "xlErrors": XlSpecialCellsValueFromString = xlErrors
    End Select
End Function

Function XlSpecialCellsValueToString(value As XlSpecialCellsValue) As String
    Select Case value
        Case xlNumbers: XlSpecialCellsValueToString = "xlNumbers"
        Case xlTextValues: XlSpecialCellsValueToString = "xlTextValues"
        Case xlLogical: XlSpecialCellsValueToString = "xlLogical"
        Case xlErrors: XlSpecialCellsValueToString = "xlErrors"
    End Select
End Function
