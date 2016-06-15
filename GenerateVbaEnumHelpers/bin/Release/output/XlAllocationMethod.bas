Attribute VB_Name = "wXlAllocationMethod"
Function XlAllocationMethodFromString(value As String) As XlAllocationMethod
    If IsNumeric(value) Then
        XlAllocationMethodFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlEqualAllocation": XlAllocationMethodFromString = xlEqualAllocation
        Case "xlWeightedAllocation": XlAllocationMethodFromString = xlWeightedAllocation
    End Select
End Function

Function XlAllocationMethodToString(value As XlAllocationMethod) As String
    Select Case value
        Case xlEqualAllocation: XlAllocationMethodToString = "xlEqualAllocation"
        Case xlWeightedAllocation: XlAllocationMethodToString = "xlWeightedAllocation"
    End Select
End Function
