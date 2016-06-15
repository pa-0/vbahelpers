Attribute VB_Name = "wWdMailMergeActiveRecord"
Function WdMailMergeActiveRecordFromString(value As String) As WdMailMergeActiveRecord
    If IsNumeric(value) Then
        WdMailMergeActiveRecordFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdPreviousDataSourceRecord": WdMailMergeActiveRecordFromString = wdPreviousDataSourceRecord
        Case "wdNextDataSourceRecord": WdMailMergeActiveRecordFromString = wdNextDataSourceRecord
        Case "wdLastDataSourceRecord": WdMailMergeActiveRecordFromString = wdLastDataSourceRecord
        Case "wdFirstDataSourceRecord": WdMailMergeActiveRecordFromString = wdFirstDataSourceRecord
        Case "wdLastRecord": WdMailMergeActiveRecordFromString = wdLastRecord
        Case "wdFirstRecord": WdMailMergeActiveRecordFromString = wdFirstRecord
        Case "wdPreviousRecord": WdMailMergeActiveRecordFromString = wdPreviousRecord
        Case "wdNextRecord": WdMailMergeActiveRecordFromString = wdNextRecord
        Case "wdNoActiveRecord": WdMailMergeActiveRecordFromString = wdNoActiveRecord
    End Select
End Function

Function WdMailMergeActiveRecordToString(value As WdMailMergeActiveRecord) As String
    Select Case value
        Case wdPreviousDataSourceRecord: WdMailMergeActiveRecordToString = "wdPreviousDataSourceRecord"
        Case wdNextDataSourceRecord: WdMailMergeActiveRecordToString = "wdNextDataSourceRecord"
        Case wdLastDataSourceRecord: WdMailMergeActiveRecordToString = "wdLastDataSourceRecord"
        Case wdFirstDataSourceRecord: WdMailMergeActiveRecordToString = "wdFirstDataSourceRecord"
        Case wdLastRecord: WdMailMergeActiveRecordToString = "wdLastRecord"
        Case wdFirstRecord: WdMailMergeActiveRecordToString = "wdFirstRecord"
        Case wdPreviousRecord: WdMailMergeActiveRecordToString = "wdPreviousRecord"
        Case wdNextRecord: WdMailMergeActiveRecordToString = "wdNextRecord"
        Case wdNoActiveRecord: WdMailMergeActiveRecordToString = "wdNoActiveRecord"
    End Select
End Function
