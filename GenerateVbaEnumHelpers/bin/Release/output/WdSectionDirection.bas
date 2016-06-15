Attribute VB_Name = "wWdSectionDirection"
Function WdSectionDirectionFromString(value As String) As WdSectionDirection
    If IsNumeric(value) Then
        WdSectionDirectionFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdSectionDirectionRtl": WdSectionDirectionFromString = wdSectionDirectionRtl
        Case "wdSectionDirectionLtr": WdSectionDirectionFromString = wdSectionDirectionLtr
    End Select
End Function

Function WdSectionDirectionToString(value As WdSectionDirection) As String
    Select Case value
        Case wdSectionDirectionRtl: WdSectionDirectionToString = "wdSectionDirectionRtl"
        Case wdSectionDirectionLtr: WdSectionDirectionToString = "wdSectionDirectionLtr"
    End Select
End Function
