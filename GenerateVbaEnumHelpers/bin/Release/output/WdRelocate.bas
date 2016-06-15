Attribute VB_Name = "wWdRelocate"
Function WdRelocateFromString(value As String) As WdRelocate
    If IsNumeric(value) Then
        WdRelocateFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdRelocateUp": WdRelocateFromString = wdRelocateUp
        Case "wdRelocateDown": WdRelocateFromString = wdRelocateDown
    End Select
End Function

Function WdRelocateToString(value As WdRelocate) As String
    Select Case value
        Case wdRelocateUp: WdRelocateToString = "wdRelocateUp"
        Case wdRelocateDown: WdRelocateToString = "wdRelocateDown"
    End Select
End Function
