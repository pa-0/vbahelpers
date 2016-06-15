Attribute VB_Name = "wWdCompareDestination"
Function WdCompareDestinationFromString(value As String) As WdCompareDestination
    If IsNumeric(value) Then
        WdCompareDestinationFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdCompareDestinationOriginal": WdCompareDestinationFromString = wdCompareDestinationOriginal
        Case "wdCompareDestinationRevised": WdCompareDestinationFromString = wdCompareDestinationRevised
        Case "wdCompareDestinationNew": WdCompareDestinationFromString = wdCompareDestinationNew
    End Select
End Function

Function WdCompareDestinationToString(value As WdCompareDestination) As String
    Select Case value
        Case wdCompareDestinationOriginal: WdCompareDestinationToString = "wdCompareDestinationOriginal"
        Case wdCompareDestinationRevised: WdCompareDestinationToString = "wdCompareDestinationRevised"
        Case wdCompareDestinationNew: WdCompareDestinationToString = "wdCompareDestinationNew"
    End Select
End Function
