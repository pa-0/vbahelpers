Attribute VB_Name = "wWdCollapseDirection"
Function WdCollapseDirectionFromString(value As String) As WdCollapseDirection
    If IsNumeric(value) Then
        WdCollapseDirectionFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdCollapseEnd": WdCollapseDirectionFromString = wdCollapseEnd
        Case "wdCollapseStart": WdCollapseDirectionFromString = wdCollapseStart
    End Select
End Function

Function WdCollapseDirectionToString(value As WdCollapseDirection) As String
    Select Case value
        Case wdCollapseEnd: WdCollapseDirectionToString = "wdCollapseEnd"
        Case wdCollapseStart: WdCollapseDirectionToString = "wdCollapseStart"
    End Select
End Function
