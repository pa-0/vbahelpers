Attribute VB_Name = "wWdMailMergeState"
Function WdMailMergeStateFromString(value As String) As WdMailMergeState
    If IsNumeric(value) Then
        WdMailMergeStateFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdNormalDocument": WdMailMergeStateFromString = wdNormalDocument
        Case "wdMainDocumentOnly": WdMailMergeStateFromString = wdMainDocumentOnly
        Case "wdMainAndDataSource": WdMailMergeStateFromString = wdMainAndDataSource
        Case "wdMainAndHeader": WdMailMergeStateFromString = wdMainAndHeader
        Case "wdMainAndSourceAndHeader": WdMailMergeStateFromString = wdMainAndSourceAndHeader
        Case "wdDataSource": WdMailMergeStateFromString = wdDataSource
    End Select
End Function

Function WdMailMergeStateToString(value As WdMailMergeState) As String
    Select Case value
        Case wdNormalDocument: WdMailMergeStateToString = "wdNormalDocument"
        Case wdMainDocumentOnly: WdMailMergeStateToString = "wdMainDocumentOnly"
        Case wdMainAndDataSource: WdMailMergeStateToString = "wdMainAndDataSource"
        Case wdMainAndHeader: WdMailMergeStateToString = "wdMainAndHeader"
        Case wdMainAndSourceAndHeader: WdMailMergeStateToString = "wdMainAndSourceAndHeader"
        Case wdDataSource: WdMailMergeStateToString = "wdDataSource"
    End Select
End Function
