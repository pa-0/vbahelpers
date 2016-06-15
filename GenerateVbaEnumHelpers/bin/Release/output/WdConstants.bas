Attribute VB_Name = "wWdConstants"
Function WdConstantsFromString(value As String) As WdConstants
    If IsNumeric(value) Then
        WdConstantsFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdAutoPosition": WdConstantsFromString = wdAutoPosition
        Case "wdFirst": WdConstantsFromString = wdFirst
        Case "wdToggle": WdConstantsFromString = wdToggle
        Case "wdUndefined": WdConstantsFromString = wdUndefined
        Case "wdForward": WdConstantsFromString = wdForward
        Case "wdCreatorCode": WdConstantsFromString = wdCreatorCode
        Case "wdBackward": WdConstantsFromString = wdBackward
    End Select
End Function

Function WdConstantsToString(value As WdConstants) As String
    Select Case value
        Case wdAutoPosition: WdConstantsToString = "wdAutoPosition"
        Case wdFirst: WdConstantsToString = "wdFirst"
        Case wdToggle: WdConstantsToString = "wdToggle"
        Case wdUndefined: WdConstantsToString = "wdUndefined"
        Case wdForward: WdConstantsToString = "wdForward"
        Case wdCreatorCode: WdConstantsToString = "wdCreatorCode"
        Case wdBackward: WdConstantsToString = "wdBackward"
    End Select
End Function
