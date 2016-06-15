Attribute VB_Name = "wPbCollapseDirection"
Function PbCollapseDirectionFromString(value As String) As PbCollapseDirection
    If IsNumeric(value) Then
        PbCollapseDirectionFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbCollapseStart": PbCollapseDirectionFromString = pbCollapseStart
        Case "pbCollapseEnd": PbCollapseDirectionFromString = pbCollapseEnd
    End Select
End Function

Function PbCollapseDirectionToString(value As PbCollapseDirection) As String
    Select Case value
        Case pbCollapseStart: PbCollapseDirectionToString = "pbCollapseStart"
        Case pbCollapseEnd: PbCollapseDirectionToString = "pbCollapseEnd"
    End Select
End Function
