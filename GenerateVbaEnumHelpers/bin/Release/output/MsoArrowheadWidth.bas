Attribute VB_Name = "wMsoArrowheadWidth"
Function MsoArrowheadWidthFromString(value As String) As MsoArrowheadWidth
    If IsNumeric(value) Then
        MsoArrowheadWidthFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoArrowheadNarrow": MsoArrowheadWidthFromString = msoArrowheadNarrow
        Case "msoArrowheadWidthMedium": MsoArrowheadWidthFromString = msoArrowheadWidthMedium
        Case "msoArrowheadWide": MsoArrowheadWidthFromString = msoArrowheadWide
        Case "msoArrowheadWidthMixed": MsoArrowheadWidthFromString = msoArrowheadWidthMixed
    End Select
End Function

Function MsoArrowheadWidthToString(value As MsoArrowheadWidth) As String
    Select Case value
        Case msoArrowheadNarrow: MsoArrowheadWidthToString = "msoArrowheadNarrow"
        Case msoArrowheadWidthMedium: MsoArrowheadWidthToString = "msoArrowheadWidthMedium"
        Case msoArrowheadWide: MsoArrowheadWidthToString = "msoArrowheadWide"
        Case msoArrowheadWidthMixed: MsoArrowheadWidthToString = "msoArrowheadWidthMixed"
    End Select
End Function
