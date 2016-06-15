Attribute VB_Name = "wMsoBarRow"
Function MsoBarRowFromString(value As String) As MsoBarRow
    If IsNumeric(value) Then
        MsoBarRowFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoBarRowFirst": MsoBarRowFromString = msoBarRowFirst
        Case "msoBarRowLast": MsoBarRowFromString = msoBarRowLast
    End Select
End Function

Function MsoBarRowToString(value As MsoBarRow) As String
    Select Case value
        Case msoBarRowFirst: MsoBarRowToString = "msoBarRowFirst"
        Case msoBarRowLast: MsoBarRowToString = "msoBarRowLast"
    End Select
End Function
