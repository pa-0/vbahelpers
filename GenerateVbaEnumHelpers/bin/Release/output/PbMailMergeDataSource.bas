Attribute VB_Name = "wPbMailMergeDataSource"
Function PbMailMergeDataSourceFromString(value As String) As PbMailMergeDataSource
    If IsNumeric(value) Then
        PbMailMergeDataSourceFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbMergeInfoFromODSO": PbMailMergeDataSourceFromString = pbMergeInfoFromODSO
        Case "pbMergeInfoSubODSO": PbMailMergeDataSourceFromString = pbMergeInfoSubODSO
    End Select
End Function

Function PbMailMergeDataSourceToString(value As PbMailMergeDataSource) As String
    Select Case value
        Case pbMergeInfoFromODSO: PbMailMergeDataSourceToString = "pbMergeInfoFromODSO"
        Case pbMergeInfoSubODSO: PbMailMergeDataSourceToString = "pbMergeInfoSubODSO"
    End Select
End Function
