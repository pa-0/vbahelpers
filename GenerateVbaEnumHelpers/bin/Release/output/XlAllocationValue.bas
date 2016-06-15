Attribute VB_Name = "wXlAllocationValue"
Function XlAllocationValueFromString(value As String) As XlAllocationValue
    If IsNumeric(value) Then
        XlAllocationValueFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlAllocateValue": XlAllocationValueFromString = xlAllocateValue
        Case "xlAllocateIncrement": XlAllocationValueFromString = xlAllocateIncrement
    End Select
End Function

Function XlAllocationValueToString(value As XlAllocationValue) As String
    Select Case value
        Case xlAllocateValue: XlAllocationValueToString = "xlAllocateValue"
        Case xlAllocateIncrement: XlAllocationValueToString = "xlAllocateIncrement"
    End Select
End Function
