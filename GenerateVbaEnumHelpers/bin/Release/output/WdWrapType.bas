Attribute VB_Name = "wWdWrapType"
Function WdWrapTypeFromString(value As String) As WdWrapType
    If IsNumeric(value) Then
        WdWrapTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdWrapSquare": WdWrapTypeFromString = wdWrapSquare
        Case "wdWrapTight": WdWrapTypeFromString = wdWrapTight
        Case "wdWrapThrough": WdWrapTypeFromString = wdWrapThrough
        Case "wdWrapNone": WdWrapTypeFromString = wdWrapNone
        Case "wdWrapFront": WdWrapTypeFromString = wdWrapFront
        Case "wdWrapTopBottom": WdWrapTypeFromString = wdWrapTopBottom
        Case "wdWrapBehind": WdWrapTypeFromString = wdWrapBehind
        Case "wdWrapInline": WdWrapTypeFromString = wdWrapInline
    End Select
End Function

Function WdWrapTypeToString(value As WdWrapType) As String
    Select Case value
        Case wdWrapSquare: WdWrapTypeToString = "wdWrapSquare"
        Case wdWrapTight: WdWrapTypeToString = "wdWrapTight"
        Case wdWrapThrough: WdWrapTypeToString = "wdWrapThrough"
        Case wdWrapNone: WdWrapTypeToString = "wdWrapNone"
        Case wdWrapFront: WdWrapTypeToString = "wdWrapFront"
        Case wdWrapTopBottom: WdWrapTypeToString = "wdWrapTopBottom"
        Case wdWrapBehind: WdWrapTypeToString = "wdWrapBehind"
        Case wdWrapInline: WdWrapTypeToString = "wdWrapInline"
    End Select
End Function
