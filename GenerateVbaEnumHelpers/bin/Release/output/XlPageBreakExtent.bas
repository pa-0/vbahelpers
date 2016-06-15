Attribute VB_Name = "wXlPageBreakExtent"
Function XlPageBreakExtentFromString(value As String) As XlPageBreakExtent
    If IsNumeric(value) Then
        XlPageBreakExtentFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlPageBreakFull": XlPageBreakExtentFromString = xlPageBreakFull
        Case "xlPageBreakPartial": XlPageBreakExtentFromString = xlPageBreakPartial
    End Select
End Function

Function XlPageBreakExtentToString(value As XlPageBreakExtent) As String
    Select Case value
        Case xlPageBreakFull: XlPageBreakExtentToString = "xlPageBreakFull"
        Case xlPageBreakPartial: XlPageBreakExtentToString = "xlPageBreakPartial"
    End Select
End Function
