Attribute VB_Name = "wWdFlowDirection"
Function WdFlowDirectionFromString(value As String) As WdFlowDirection
    If IsNumeric(value) Then
        WdFlowDirectionFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdFlowLtr": WdFlowDirectionFromString = wdFlowLtr
        Case "wdFlowRtl": WdFlowDirectionFromString = wdFlowRtl
    End Select
End Function

Function WdFlowDirectionToString(value As WdFlowDirection) As String
    Select Case value
        Case wdFlowLtr: WdFlowDirectionToString = "wdFlowLtr"
        Case wdFlowRtl: WdFlowDirectionToString = "wdFlowRtl"
    End Select
End Function
