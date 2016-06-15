Attribute VB_Name = "wWdJustificationMode"
Function WdJustificationModeFromString(value As String) As WdJustificationMode
    If IsNumeric(value) Then
        WdJustificationModeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdJustificationModeExpand": WdJustificationModeFromString = wdJustificationModeExpand
        Case "wdJustificationModeCompress": WdJustificationModeFromString = wdJustificationModeCompress
        Case "wdJustificationModeCompressKana": WdJustificationModeFromString = wdJustificationModeCompressKana
    End Select
End Function

Function WdJustificationModeToString(value As WdJustificationMode) As String
    Select Case value
        Case wdJustificationModeExpand: WdJustificationModeToString = "wdJustificationModeExpand"
        Case wdJustificationModeCompress: WdJustificationModeToString = "wdJustificationModeCompress"
        Case wdJustificationModeCompressKana: WdJustificationModeToString = "wdJustificationModeCompressKana"
    End Select
End Function
