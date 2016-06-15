Attribute VB_Name = "wMsoSyncVersionType"
Function MsoSyncVersionTypeFromString(value As String) As MsoSyncVersionType
    If IsNumeric(value) Then
        MsoSyncVersionTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoSyncVersionLastViewed": MsoSyncVersionTypeFromString = msoSyncVersionLastViewed
        Case "msoSyncVersionServer": MsoSyncVersionTypeFromString = msoSyncVersionServer
    End Select
End Function

Function MsoSyncVersionTypeToString(value As MsoSyncVersionType) As String
    Select Case value
        Case msoSyncVersionLastViewed: MsoSyncVersionTypeToString = "msoSyncVersionLastViewed"
        Case msoSyncVersionServer: MsoSyncVersionTypeToString = "msoSyncVersionServer"
    End Select
End Function
