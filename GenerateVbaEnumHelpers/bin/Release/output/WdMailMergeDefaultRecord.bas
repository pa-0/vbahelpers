Attribute VB_Name = "wWdMailMergeDefaultRecord"
Function WdMailMergeDefaultRecordFromString(value As String) As WdMailMergeDefaultRecord
    If IsNumeric(value) Then
        WdMailMergeDefaultRecordFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdDefaultFirstRecord": WdMailMergeDefaultRecordFromString = wdDefaultFirstRecord
        Case "wdDefaultLastRecord": WdMailMergeDefaultRecordFromString = wdDefaultLastRecord
    End Select
End Function

Function WdMailMergeDefaultRecordToString(value As WdMailMergeDefaultRecord) As String
    Select Case value
        Case wdDefaultFirstRecord: WdMailMergeDefaultRecordToString = "wdDefaultFirstRecord"
        Case wdDefaultLastRecord: WdMailMergeDefaultRecordToString = "wdDefaultLastRecord"
    End Select
End Function
