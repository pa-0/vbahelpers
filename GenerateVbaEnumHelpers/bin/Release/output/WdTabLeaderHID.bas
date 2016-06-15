Attribute VB_Name = "wWdTabLeaderHID"
Function WdTabLeaderHIDFromString(value As String) As WdTabLeaderHID
    If IsNumeric(value) Then
        WdTabLeaderHIDFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "emptyenum": WdTabLeaderHIDFromString = emptyenum
    End Select
End Function

Function WdTabLeaderHIDToString(value As WdTabLeaderHID) As String
    Select Case value
        Case emptyenum: WdTabLeaderHIDToString = "emptyenum"
    End Select
End Function
