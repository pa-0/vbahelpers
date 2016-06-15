Attribute VB_Name = "wMsoSyncAvailableType"
Function MsoSyncAvailableTypeFromString(value As String) As MsoSyncAvailableType
    If IsNumeric(value) Then
        MsoSyncAvailableTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoSyncAvailableNone": MsoSyncAvailableTypeFromString = msoSyncAvailableNone
        Case "msoSyncAvailableOffline": MsoSyncAvailableTypeFromString = msoSyncAvailableOffline
        Case "msoSyncAvailableAnywhere": MsoSyncAvailableTypeFromString = msoSyncAvailableAnywhere
    End Select
End Function

Function MsoSyncAvailableTypeToString(value As MsoSyncAvailableType) As String
    Select Case value
        Case msoSyncAvailableNone: MsoSyncAvailableTypeToString = "msoSyncAvailableNone"
        Case msoSyncAvailableOffline: MsoSyncAvailableTypeToString = "msoSyncAvailableOffline"
        Case msoSyncAvailableAnywhere: MsoSyncAvailableTypeToString = "msoSyncAvailableAnywhere"
    End Select
End Function
