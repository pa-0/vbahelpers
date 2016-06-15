Attribute VB_Name = "wMsoSyncCompareType"
Function MsoSyncCompareTypeFromString(value As String) As MsoSyncCompareType
    If IsNumeric(value) Then
        MsoSyncCompareTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoSyncCompareAndMerge": MsoSyncCompareTypeFromString = msoSyncCompareAndMerge
        Case "msoSyncCompareSideBySide": MsoSyncCompareTypeFromString = msoSyncCompareSideBySide
    End Select
End Function

Function MsoSyncCompareTypeToString(value As MsoSyncCompareType) As String
    Select Case value
        Case msoSyncCompareAndMerge: MsoSyncCompareTypeToString = "msoSyncCompareAndMerge"
        Case msoSyncCompareSideBySide: MsoSyncCompareTypeToString = "msoSyncCompareSideBySide"
    End Select
End Function
