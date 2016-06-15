Attribute VB_Name = "wMsoMoveRow"
Function MsoMoveRowFromString(value As String) As MsoMoveRow
    If IsNumeric(value) Then
        MsoMoveRowFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoMoveRowFirst": MsoMoveRowFromString = msoMoveRowFirst
        Case "msoMoveRowPrev": MsoMoveRowFromString = msoMoveRowPrev
        Case "msoMoveRowNext": MsoMoveRowFromString = msoMoveRowNext
        Case "msoMoveRowNbr": MsoMoveRowFromString = msoMoveRowNbr
    End Select
End Function

Function MsoMoveRowToString(value As MsoMoveRow) As String
    Select Case value
        Case msoMoveRowFirst: MsoMoveRowToString = "msoMoveRowFirst"
        Case msoMoveRowPrev: MsoMoveRowToString = "msoMoveRowPrev"
        Case msoMoveRowNext: MsoMoveRowToString = "msoMoveRowNext"
        Case msoMoveRowNbr: MsoMoveRowToString = "msoMoveRowNbr"
    End Select
End Function
