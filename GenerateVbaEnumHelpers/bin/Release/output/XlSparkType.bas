Attribute VB_Name = "wXlSparkType"
Function XlSparkTypeFromString(value As String) As XlSparkType
    If IsNumeric(value) Then
        XlSparkTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlSparkLine": XlSparkTypeFromString = xlSparkLine
        Case "xlSparkColumn": XlSparkTypeFromString = xlSparkColumn
        Case "xlSparkColumnStacked100": XlSparkTypeFromString = xlSparkColumnStacked100
    End Select
End Function

Function XlSparkTypeToString(value As XlSparkType) As String
    Select Case value
        Case xlSparkLine: XlSparkTypeToString = "xlSparkLine"
        Case xlSparkColumn: XlSparkTypeToString = "xlSparkColumn"
        Case xlSparkColumnStacked100: XlSparkTypeToString = "xlSparkColumnStacked100"
    End Select
End Function
