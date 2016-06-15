Attribute VB_Name = "wWdOMathVertAlignType"
Function WdOMathVertAlignTypeFromString(value As String) As WdOMathVertAlignType
    If IsNumeric(value) Then
        WdOMathVertAlignTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdOMathVertAlignCenter": WdOMathVertAlignTypeFromString = wdOMathVertAlignCenter
        Case "wdOMathVertAlignTop": WdOMathVertAlignTypeFromString = wdOMathVertAlignTop
        Case "wdOMathVertAlignBottom": WdOMathVertAlignTypeFromString = wdOMathVertAlignBottom
    End Select
End Function

Function WdOMathVertAlignTypeToString(value As WdOMathVertAlignType) As String
    Select Case value
        Case wdOMathVertAlignCenter: WdOMathVertAlignTypeToString = "wdOMathVertAlignCenter"
        Case wdOMathVertAlignTop: WdOMathVertAlignTypeToString = "wdOMathVertAlignTop"
        Case wdOMathVertAlignBottom: WdOMathVertAlignTypeToString = "wdOMathVertAlignBottom"
    End Select
End Function
