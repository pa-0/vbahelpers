Attribute VB_Name = "wWdOMathHorizAlignType"
Function WdOMathHorizAlignTypeFromString(value As String) As WdOMathHorizAlignType
    If IsNumeric(value) Then
        WdOMathHorizAlignTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdOMathHorizAlignCenter": WdOMathHorizAlignTypeFromString = wdOMathHorizAlignCenter
        Case "wdOMathHorizAlignLeft": WdOMathHorizAlignTypeFromString = wdOMathHorizAlignLeft
        Case "wdOMathHorizAlignRight": WdOMathHorizAlignTypeFromString = wdOMathHorizAlignRight
    End Select
End Function

Function WdOMathHorizAlignTypeToString(value As WdOMathHorizAlignType) As String
    Select Case value
        Case wdOMathHorizAlignCenter: WdOMathHorizAlignTypeToString = "wdOMathHorizAlignCenter"
        Case wdOMathHorizAlignLeft: WdOMathHorizAlignTypeToString = "wdOMathHorizAlignLeft"
        Case wdOMathHorizAlignRight: WdOMathHorizAlignTypeToString = "wdOMathHorizAlignRight"
    End Select
End Function
