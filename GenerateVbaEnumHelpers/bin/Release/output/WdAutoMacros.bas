Attribute VB_Name = "wWdAutoMacros"
Function WdAutoMacrosFromString(value As String) As WdAutoMacros
    If IsNumeric(value) Then
        WdAutoMacrosFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdAutoExec": WdAutoMacrosFromString = wdAutoExec
        Case "wdAutoNew": WdAutoMacrosFromString = wdAutoNew
        Case "wdAutoOpen": WdAutoMacrosFromString = wdAutoOpen
        Case "wdAutoClose": WdAutoMacrosFromString = wdAutoClose
        Case "wdAutoExit": WdAutoMacrosFromString = wdAutoExit
        Case "wdAutoSync": WdAutoMacrosFromString = wdAutoSync
    End Select
End Function

Function WdAutoMacrosToString(value As WdAutoMacros) As String
    Select Case value
        Case wdAutoExec: WdAutoMacrosToString = "wdAutoExec"
        Case wdAutoNew: WdAutoMacrosToString = "wdAutoNew"
        Case wdAutoOpen: WdAutoMacrosToString = "wdAutoOpen"
        Case wdAutoClose: WdAutoMacrosToString = "wdAutoClose"
        Case wdAutoExit: WdAutoMacrosToString = "wdAutoExit"
        Case wdAutoSync: WdAutoMacrosToString = "wdAutoSync"
    End Select
End Function
