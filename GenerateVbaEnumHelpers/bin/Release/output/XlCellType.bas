Attribute VB_Name = "wXlCellType"
Function XlCellTypeFromString(value As String) As XlCellType
    If IsNumeric(value) Then
        XlCellTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlCellTypeConstants": XlCellTypeFromString = xlCellTypeConstants
        Case "xlCellTypeBlanks": XlCellTypeFromString = xlCellTypeBlanks
        Case "xlCellTypeLastCell": XlCellTypeFromString = xlCellTypeLastCell
        Case "xlCellTypeVisible": XlCellTypeFromString = xlCellTypeVisible
        Case "xlCellTypeSameValidation": XlCellTypeFromString = xlCellTypeSameValidation
        Case "xlCellTypeAllValidation": XlCellTypeFromString = xlCellTypeAllValidation
        Case "xlCellTypeSameFormatConditions": XlCellTypeFromString = xlCellTypeSameFormatConditions
        Case "xlCellTypeAllFormatConditions": XlCellTypeFromString = xlCellTypeAllFormatConditions
        Case "xlCellTypeComments": XlCellTypeFromString = xlCellTypeComments
        Case "xlCellTypeFormulas": XlCellTypeFromString = xlCellTypeFormulas
    End Select
End Function

Function XlCellTypeToString(value As XlCellType) As String
    Select Case value
        Case xlCellTypeConstants: XlCellTypeToString = "xlCellTypeConstants"
        Case xlCellTypeBlanks: XlCellTypeToString = "xlCellTypeBlanks"
        Case xlCellTypeLastCell: XlCellTypeToString = "xlCellTypeLastCell"
        Case xlCellTypeVisible: XlCellTypeToString = "xlCellTypeVisible"
        Case xlCellTypeSameValidation: XlCellTypeToString = "xlCellTypeSameValidation"
        Case xlCellTypeAllValidation: XlCellTypeToString = "xlCellTypeAllValidation"
        Case xlCellTypeSameFormatConditions: XlCellTypeToString = "xlCellTypeSameFormatConditions"
        Case xlCellTypeAllFormatConditions: XlCellTypeToString = "xlCellTypeAllFormatConditions"
        Case xlCellTypeComments: XlCellTypeToString = "xlCellTypeComments"
        Case xlCellTypeFormulas: XlCellTypeToString = "xlCellTypeFormulas"
    End Select
End Function
