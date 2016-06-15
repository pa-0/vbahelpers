Attribute VB_Name = "wWdTableDirection"
Function WdTableDirectionFromString(value As String) As WdTableDirection
    If IsNumeric(value) Then
        WdTableDirectionFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdTableDirectionRtl": WdTableDirectionFromString = wdTableDirectionRtl
        Case "wdTableDirectionLtr": WdTableDirectionFromString = wdTableDirectionLtr
    End Select
End Function

Function WdTableDirectionToString(value As WdTableDirection) As String
    Select Case value
        Case wdTableDirectionRtl: WdTableDirectionToString = "wdTableDirectionRtl"
        Case wdTableDirectionLtr: WdTableDirectionToString = "wdTableDirectionLtr"
    End Select
End Function
