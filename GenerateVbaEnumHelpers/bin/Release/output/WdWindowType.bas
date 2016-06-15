Attribute VB_Name = "wWdWindowType"
Function WdWindowTypeFromString(value As String) As WdWindowType
    If IsNumeric(value) Then
        WdWindowTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdWindowDocument": WdWindowTypeFromString = wdWindowDocument
        Case "wdWindowTemplate": WdWindowTypeFromString = wdWindowTemplate
    End Select
End Function

Function WdWindowTypeToString(value As WdWindowType) As String
    Select Case value
        Case wdWindowDocument: WdWindowTypeToString = "wdWindowDocument"
        Case wdWindowTemplate: WdWindowTypeToString = "wdWindowTemplate"
    End Select
End Function
