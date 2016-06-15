Attribute VB_Name = "wMsoArrowheadLength"
Function MsoArrowheadLengthFromString(value As String) As MsoArrowheadLength
    If IsNumeric(value) Then
        MsoArrowheadLengthFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoArrowheadShort": MsoArrowheadLengthFromString = msoArrowheadShort
        Case "msoArrowheadLengthMedium": MsoArrowheadLengthFromString = msoArrowheadLengthMedium
        Case "msoArrowheadLong": MsoArrowheadLengthFromString = msoArrowheadLong
        Case "msoArrowheadLengthMixed": MsoArrowheadLengthFromString = msoArrowheadLengthMixed
    End Select
End Function

Function MsoArrowheadLengthToString(value As MsoArrowheadLength) As String
    Select Case value
        Case msoArrowheadShort: MsoArrowheadLengthToString = "msoArrowheadShort"
        Case msoArrowheadLengthMedium: MsoArrowheadLengthToString = "msoArrowheadLengthMedium"
        Case msoArrowheadLong: MsoArrowheadLengthToString = "msoArrowheadLong"
        Case msoArrowheadLengthMixed: MsoArrowheadLengthToString = "msoArrowheadLengthMixed"
    End Select
End Function
