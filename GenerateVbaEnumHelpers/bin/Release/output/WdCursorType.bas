Attribute VB_Name = "wWdCursorType"
Function WdCursorTypeFromString(value As String) As WdCursorType
    If IsNumeric(value) Then
        WdCursorTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdCursorWait": WdCursorTypeFromString = wdCursorWait
        Case "wdCursorIBeam": WdCursorTypeFromString = wdCursorIBeam
        Case "wdCursorNormal": WdCursorTypeFromString = wdCursorNormal
        Case "wdCursorNorthwestArrow": WdCursorTypeFromString = wdCursorNorthwestArrow
    End Select
End Function

Function WdCursorTypeToString(value As WdCursorType) As String
    Select Case value
        Case wdCursorWait: WdCursorTypeToString = "wdCursorWait"
        Case wdCursorIBeam: WdCursorTypeToString = "wdCursorIBeam"
        Case wdCursorNormal: WdCursorTypeToString = "wdCursorNormal"
        Case wdCursorNorthwestArrow: WdCursorTypeToString = "wdCursorNorthwestArrow"
    End Select
End Function
