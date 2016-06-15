Attribute VB_Name = "wWdAutoVersions"
Function WdAutoVersionsFromString(value As String) As WdAutoVersions
    If IsNumeric(value) Then
        WdAutoVersionsFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdAutoVersionOff": WdAutoVersionsFromString = wdAutoVersionOff
        Case "wdAutoVersionOnClose": WdAutoVersionsFromString = wdAutoVersionOnClose
    End Select
End Function

Function WdAutoVersionsToString(value As WdAutoVersions) As String
    Select Case value
        Case wdAutoVersionOff: WdAutoVersionsToString = "wdAutoVersionOff"
        Case wdAutoVersionOnClose: WdAutoVersionsToString = "wdAutoVersionOnClose"
    End Select
End Function
