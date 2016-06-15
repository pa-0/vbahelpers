Attribute VB_Name = "wWdOLEType"
Function WdOLETypeFromString(value As String) As WdOLEType
    If IsNumeric(value) Then
        WdOLETypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdOLELink": WdOLETypeFromString = wdOLELink
        Case "wdOLEEmbed": WdOLETypeFromString = wdOLEEmbed
        Case "wdOLEControl": WdOLETypeFromString = wdOLEControl
    End Select
End Function

Function WdOLETypeToString(value As WdOLEType) As String
    Select Case value
        Case wdOLELink: WdOLETypeToString = "wdOLELink"
        Case wdOLEEmbed: WdOLETypeToString = "wdOLEEmbed"
        Case wdOLEControl: WdOLETypeToString = "wdOLEControl"
    End Select
End Function
