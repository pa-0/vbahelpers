Attribute VB_Name = "wWdOMathJc"
Function WdOMathJcFromString(value As String) As WdOMathJc
    If IsNumeric(value) Then
        WdOMathJcFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdOMathJcCenterGroup": WdOMathJcFromString = wdOMathJcCenterGroup
        Case "wdOMathJcCenter": WdOMathJcFromString = wdOMathJcCenter
        Case "wdOMathJcLeft": WdOMathJcFromString = wdOMathJcLeft
        Case "wdOMathJcRight": WdOMathJcFromString = wdOMathJcRight
        Case "wdOMathJcInline": WdOMathJcFromString = wdOMathJcInline
    End Select
End Function

Function WdOMathJcToString(value As WdOMathJc) As String
    Select Case value
        Case wdOMathJcCenterGroup: WdOMathJcToString = "wdOMathJcCenterGroup"
        Case wdOMathJcCenter: WdOMathJcToString = "wdOMathJcCenter"
        Case wdOMathJcLeft: WdOMathJcToString = "wdOMathJcLeft"
        Case wdOMathJcRight: WdOMathJcToString = "wdOMathJcRight"
        Case wdOMathJcInline: WdOMathJcToString = "wdOMathJcInline"
    End Select
End Function
