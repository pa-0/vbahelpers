Attribute VB_Name = "wMsoSyncEventType"
Function MsoSyncEventTypeFromString(value As String) As MsoSyncEventType
    If IsNumeric(value) Then
        MsoSyncEventTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoSyncEventDownloadInitiated": MsoSyncEventTypeFromString = msoSyncEventDownloadInitiated
        Case "msoSyncEventDownloadSucceeded": MsoSyncEventTypeFromString = msoSyncEventDownloadSucceeded
        Case "msoSyncEventDownloadFailed": MsoSyncEventTypeFromString = msoSyncEventDownloadFailed
        Case "msoSyncEventUploadInitiated": MsoSyncEventTypeFromString = msoSyncEventUploadInitiated
        Case "msoSyncEventUploadSucceeded": MsoSyncEventTypeFromString = msoSyncEventUploadSucceeded
        Case "msoSyncEventUploadFailed": MsoSyncEventTypeFromString = msoSyncEventUploadFailed
        Case "msoSyncEventDownloadNoChange": MsoSyncEventTypeFromString = msoSyncEventDownloadNoChange
        Case "msoSyncEventOffline": MsoSyncEventTypeFromString = msoSyncEventOffline
    End Select
End Function

Function MsoSyncEventTypeToString(value As MsoSyncEventType) As String
    Select Case value
        Case msoSyncEventDownloadInitiated: MsoSyncEventTypeToString = "msoSyncEventDownloadInitiated"
        Case msoSyncEventDownloadSucceeded: MsoSyncEventTypeToString = "msoSyncEventDownloadSucceeded"
        Case msoSyncEventDownloadFailed: MsoSyncEventTypeToString = "msoSyncEventDownloadFailed"
        Case msoSyncEventUploadInitiated: MsoSyncEventTypeToString = "msoSyncEventUploadInitiated"
        Case msoSyncEventUploadSucceeded: MsoSyncEventTypeToString = "msoSyncEventUploadSucceeded"
        Case msoSyncEventUploadFailed: MsoSyncEventTypeToString = "msoSyncEventUploadFailed"
        Case msoSyncEventDownloadNoChange: MsoSyncEventTypeToString = "msoSyncEventDownloadNoChange"
        Case msoSyncEventOffline: MsoSyncEventTypeToString = "msoSyncEventOffline"
    End Select
End Function
