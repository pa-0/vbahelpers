Attribute VB_Name = "wXlCalcFor"
Function XlCalcForFromString(value As String) As XlCalcFor
    If IsNumeric(value) Then
        XlCalcForFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlAllValues": XlCalcForFromString = xlAllValues
        Case "xlRowGroups": XlCalcForFromString = xlRowGroups
        Case "xlColGroups": XlCalcForFromString = xlColGroups
    End Select
End Function

Function XlCalcForToString(value As XlCalcFor) As String
    Select Case value
        Case xlAllValues: XlCalcForToString = "xlAllValues"
        Case xlRowGroups: XlCalcForToString = "xlRowGroups"
        Case xlColGroups: XlCalcForToString = "xlColGroups"
    End Select
End Function
