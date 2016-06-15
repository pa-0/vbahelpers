Attribute VB_Name = "wWdVisualSelection"
Function WdVisualSelectionFromString(value As String) As WdVisualSelection
    If IsNumeric(value) Then
        WdVisualSelectionFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdVisualSelectionBlock": WdVisualSelectionFromString = wdVisualSelectionBlock
        Case "wdVisualSelectionContinuous": WdVisualSelectionFromString = wdVisualSelectionContinuous
    End Select
End Function

Function WdVisualSelectionToString(value As WdVisualSelection) As String
    Select Case value
        Case wdVisualSelectionBlock: WdVisualSelectionToString = "wdVisualSelectionBlock"
        Case wdVisualSelectionContinuous: WdVisualSelectionToString = "wdVisualSelectionContinuous"
    End Select
End Function
