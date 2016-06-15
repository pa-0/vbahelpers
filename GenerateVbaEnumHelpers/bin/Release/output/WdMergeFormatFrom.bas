Attribute VB_Name = "wWdMergeFormatFrom"
Function WdMergeFormatFromFromString(value As String) As WdMergeFormatFrom
    If IsNumeric(value) Then
        WdMergeFormatFromFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdMergeFormatFromOriginal": WdMergeFormatFromFromString = wdMergeFormatFromOriginal
        Case "wdMergeFormatFromRevised": WdMergeFormatFromFromString = wdMergeFormatFromRevised
        Case "wdMergeFormatFromPrompt": WdMergeFormatFromFromString = wdMergeFormatFromPrompt
    End Select
End Function

Function WdMergeFormatFromToString(value As WdMergeFormatFrom) As String
    Select Case value
        Case wdMergeFormatFromOriginal: WdMergeFormatFromToString = "wdMergeFormatFromOriginal"
        Case wdMergeFormatFromRevised: WdMergeFormatFromToString = "wdMergeFormatFromRevised"
        Case wdMergeFormatFromPrompt: WdMergeFormatFromToString = "wdMergeFormatFromPrompt"
    End Select
End Function
