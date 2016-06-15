Attribute VB_Name = "wXlAllocation"
Function XlAllocationFromString(value As String) As XlAllocation
    If IsNumeric(value) Then
        XlAllocationFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlManualAllocation": XlAllocationFromString = xlManualAllocation
        Case "xlAutomaticAllocation": XlAllocationFromString = xlAutomaticAllocation
    End Select
End Function

Function XlAllocationToString(value As XlAllocation) As String
    Select Case value
        Case xlManualAllocation: XlAllocationToString = "xlManualAllocation"
        Case xlAutomaticAllocation: XlAllocationToString = "xlAutomaticAllocation"
    End Select
End Function
