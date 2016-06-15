Attribute VB_Name = "wPbCellDiagonalType"
Function PbCellDiagonalTypeFromString(value As String) As PbCellDiagonalType
    If IsNumeric(value) Then
        PbCellDiagonalTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbTableCellDiagonalNone": PbCellDiagonalTypeFromString = pbTableCellDiagonalNone
        Case "pbTableCellDiagonalUp": PbCellDiagonalTypeFromString = pbTableCellDiagonalUp
        Case "pbTableCellDiagonalDown": PbCellDiagonalTypeFromString = pbTableCellDiagonalDown
        Case "pbTableCellDiagonalMixed": PbCellDiagonalTypeFromString = pbTableCellDiagonalMixed
    End Select
End Function

Function PbCellDiagonalTypeToString(value As PbCellDiagonalType) As String
    Select Case value
        Case pbTableCellDiagonalNone: PbCellDiagonalTypeToString = "pbTableCellDiagonalNone"
        Case pbTableCellDiagonalUp: PbCellDiagonalTypeToString = "pbTableCellDiagonalUp"
        Case pbTableCellDiagonalDown: PbCellDiagonalTypeToString = "pbTableCellDiagonalDown"
        Case pbTableCellDiagonalMixed: PbCellDiagonalTypeToString = "pbTableCellDiagonalMixed"
    End Select
End Function
