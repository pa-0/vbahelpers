Attribute VB_Name = "wWdFieldKind"
Function WdFieldKindFromString(value As String) As WdFieldKind
    If IsNumeric(value) Then
        WdFieldKindFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdFieldKindNone": WdFieldKindFromString = wdFieldKindNone
        Case "wdFieldKindHot": WdFieldKindFromString = wdFieldKindHot
        Case "wdFieldKindWarm": WdFieldKindFromString = wdFieldKindWarm
        Case "wdFieldKindCold": WdFieldKindFromString = wdFieldKindCold
    End Select
End Function

Function WdFieldKindToString(value As WdFieldKind) As String
    Select Case value
        Case wdFieldKindNone: WdFieldKindToString = "wdFieldKindNone"
        Case wdFieldKindHot: WdFieldKindToString = "wdFieldKindHot"
        Case wdFieldKindWarm: WdFieldKindToString = "wdFieldKindWarm"
        Case wdFieldKindCold: WdFieldKindToString = "wdFieldKindCold"
    End Select
End Function
