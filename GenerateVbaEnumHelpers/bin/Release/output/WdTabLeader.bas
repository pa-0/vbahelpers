Attribute VB_Name = "wWdTabLeader"
Function WdTabLeaderFromString(value As String) As WdTabLeader
    If IsNumeric(value) Then
        WdTabLeaderFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdTabLeaderSpaces": WdTabLeaderFromString = wdTabLeaderSpaces
        Case "wdTabLeaderDots": WdTabLeaderFromString = wdTabLeaderDots
        Case "wdTabLeaderDashes": WdTabLeaderFromString = wdTabLeaderDashes
        Case "wdTabLeaderLines": WdTabLeaderFromString = wdTabLeaderLines
        Case "wdTabLeaderHeavy": WdTabLeaderFromString = wdTabLeaderHeavy
        Case "wdTabLeaderMiddleDot": WdTabLeaderFromString = wdTabLeaderMiddleDot
    End Select
End Function

Function WdTabLeaderToString(value As WdTabLeader) As String
    Select Case value
        Case wdTabLeaderSpaces: WdTabLeaderToString = "wdTabLeaderSpaces"
        Case wdTabLeaderDots: WdTabLeaderToString = "wdTabLeaderDots"
        Case wdTabLeaderDashes: WdTabLeaderToString = "wdTabLeaderDashes"
        Case wdTabLeaderLines: WdTabLeaderToString = "wdTabLeaderLines"
        Case wdTabLeaderHeavy: WdTabLeaderToString = "wdTabLeaderHeavy"
        Case wdTabLeaderMiddleDot: WdTabLeaderToString = "wdTabLeaderMiddleDot"
    End Select
End Function
