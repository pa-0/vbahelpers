Attribute VB_Name = "wWdGutterStyle"
Function WdGutterStyleFromString(value As String) As WdGutterStyle
    If IsNumeric(value) Then
        WdGutterStyleFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdGutterPosLeft": WdGutterStyleFromString = wdGutterPosLeft
        Case "wdGutterPosTop": WdGutterStyleFromString = wdGutterPosTop
        Case "wdGutterPosRight": WdGutterStyleFromString = wdGutterPosRight
    End Select
End Function

Function WdGutterStyleToString(value As WdGutterStyle) As String
    Select Case value
        Case wdGutterPosLeft: WdGutterStyleToString = "wdGutterPosLeft"
        Case wdGutterPosTop: WdGutterStyleToString = "wdGutterPosTop"
        Case wdGutterPosRight: WdGutterStyleToString = "wdGutterPosRight"
    End Select
End Function
