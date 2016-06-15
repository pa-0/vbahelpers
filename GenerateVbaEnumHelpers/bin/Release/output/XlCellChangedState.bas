Attribute VB_Name = "wXlCellChangedState"
Function XlCellChangedStateFromString(value As String) As XlCellChangedState
    If IsNumeric(value) Then
        XlCellChangedStateFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlCellNotChanged": XlCellChangedStateFromString = xlCellNotChanged
        Case "xlCellChanged": XlCellChangedStateFromString = xlCellChanged
        Case "xlCellChangeApplied": XlCellChangedStateFromString = xlCellChangeApplied
    End Select
End Function

Function XlCellChangedStateToString(value As XlCellChangedState) As String
    Select Case value
        Case xlCellNotChanged: XlCellChangedStateToString = "xlCellNotChanged"
        Case xlCellChanged: XlCellChangedStateToString = "xlCellChanged"
        Case xlCellChangeApplied: XlCellChangedStateToString = "xlCellChangeApplied"
    End Select
End Function
