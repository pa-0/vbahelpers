Attribute VB_Name = "wMsoDistributeCmd"
Function MsoDistributeCmdFromString(value As String) As MsoDistributeCmd
    If IsNumeric(value) Then
        MsoDistributeCmdFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoDistributeHorizontally": MsoDistributeCmdFromString = msoDistributeHorizontally
        Case "msoDistributeVertically": MsoDistributeCmdFromString = msoDistributeVertically
    End Select
End Function

Function MsoDistributeCmdToString(value As MsoDistributeCmd) As String
    Select Case value
        Case msoDistributeHorizontally: MsoDistributeCmdToString = "msoDistributeHorizontally"
        Case msoDistributeVertically: MsoDistributeCmdToString = "msoDistributeVertically"
    End Select
End Function
