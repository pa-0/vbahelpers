Attribute VB_Name = "wOlSelectionLocation"
Function OlSelectionLocationFromString(value As String) As OlSelectionLocation
    If IsNumeric(value) Then
        OlSelectionLocationFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olViewList": OlSelectionLocationFromString = olViewList
        Case "olToDoBarTaskList": OlSelectionLocationFromString = olToDoBarTaskList
        Case "olToDoBarAppointmentList": OlSelectionLocationFromString = olToDoBarAppointmentList
        Case "olDailyTaskList": OlSelectionLocationFromString = olDailyTaskList
        Case "olAttachmentWell": OlSelectionLocationFromString = olAttachmentWell
    End Select
End Function

Function OlSelectionLocationToString(value As OlSelectionLocation) As String
    Select Case value
        Case olViewList: OlSelectionLocationToString = "olViewList"
        Case olToDoBarTaskList: OlSelectionLocationToString = "olToDoBarTaskList"
        Case olToDoBarAppointmentList: OlSelectionLocationToString = "olToDoBarAppointmentList"
        Case olDailyTaskList: OlSelectionLocationToString = "olDailyTaskList"
        Case olAttachmentWell: OlSelectionLocationToString = "olAttachmentWell"
    End Select
End Function
