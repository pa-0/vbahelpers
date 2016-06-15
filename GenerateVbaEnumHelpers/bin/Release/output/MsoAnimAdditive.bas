Attribute VB_Name = "wMsoAnimAdditive"
Function MsoAnimAdditiveFromString(value As String) As MsoAnimAdditive
    If IsNumeric(value) Then
        MsoAnimAdditiveFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoAnimAdditiveAddBase": MsoAnimAdditiveFromString = msoAnimAdditiveAddBase
        Case "msoAnimAdditiveAddSum": MsoAnimAdditiveFromString = msoAnimAdditiveAddSum
    End Select
End Function

Function MsoAnimAdditiveToString(value As MsoAnimAdditive) As String
    Select Case value
        Case msoAnimAdditiveAddBase: MsoAnimAdditiveToString = "msoAnimAdditiveAddBase"
        Case msoAnimAdditiveAddSum: MsoAnimAdditiveToString = "msoAnimAdditiveAddSum"
    End Select
End Function
