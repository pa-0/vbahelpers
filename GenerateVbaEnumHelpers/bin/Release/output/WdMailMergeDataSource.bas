Attribute VB_Name = "wWdMailMergeDataSource"
Function WdMailMergeDataSourceFromString(value As String) As WdMailMergeDataSource
    If IsNumeric(value) Then
        WdMailMergeDataSourceFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdMergeInfoFromWord": WdMailMergeDataSourceFromString = wdMergeInfoFromWord
        Case "wdMergeInfoFromAccessDDE": WdMailMergeDataSourceFromString = wdMergeInfoFromAccessDDE
        Case "wdMergeInfoFromExcelDDE": WdMailMergeDataSourceFromString = wdMergeInfoFromExcelDDE
        Case "wdMergeInfoFromMSQueryDDE": WdMailMergeDataSourceFromString = wdMergeInfoFromMSQueryDDE
        Case "wdMergeInfoFromODBC": WdMailMergeDataSourceFromString = wdMergeInfoFromODBC
        Case "wdMergeInfoFromODSO": WdMailMergeDataSourceFromString = wdMergeInfoFromODSO
        Case "wdNoMergeInfo": WdMailMergeDataSourceFromString = wdNoMergeInfo
    End Select
End Function

Function WdMailMergeDataSourceToString(value As WdMailMergeDataSource) As String
    Select Case value
        Case wdMergeInfoFromWord: WdMailMergeDataSourceToString = "wdMergeInfoFromWord"
        Case wdMergeInfoFromAccessDDE: WdMailMergeDataSourceToString = "wdMergeInfoFromAccessDDE"
        Case wdMergeInfoFromExcelDDE: WdMailMergeDataSourceToString = "wdMergeInfoFromExcelDDE"
        Case wdMergeInfoFromMSQueryDDE: WdMailMergeDataSourceToString = "wdMergeInfoFromMSQueryDDE"
        Case wdMergeInfoFromODBC: WdMailMergeDataSourceToString = "wdMergeInfoFromODBC"
        Case wdMergeInfoFromODSO: WdMailMergeDataSourceToString = "wdMergeInfoFromODSO"
        Case wdNoMergeInfo: WdMailMergeDataSourceToString = "wdNoMergeInfo"
    End Select
End Function
