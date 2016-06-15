Attribute VB_Name = "wMsoAnimAccumulate"
Function MsoAnimAccumulateFromString(value As String) As MsoAnimAccumulate
    If IsNumeric(value) Then
        MsoAnimAccumulateFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoAnimAccumulateNone": MsoAnimAccumulateFromString = msoAnimAccumulateNone
        Case "msoAnimAccumulateAlways": MsoAnimAccumulateFromString = msoAnimAccumulateAlways
    End Select
End Function

Function MsoAnimAccumulateToString(value As MsoAnimAccumulate) As String
    Select Case value
        Case msoAnimAccumulateNone: MsoAnimAccumulateToString = "msoAnimAccumulateNone"
        Case msoAnimAccumulateAlways: MsoAnimAccumulateToString = "msoAnimAccumulateAlways"
    End Select
End Function
