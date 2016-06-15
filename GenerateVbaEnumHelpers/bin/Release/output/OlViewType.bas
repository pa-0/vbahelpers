Attribute VB_Name = "wOlViewType"
Function OlViewTypeFromString(value As String) As OlViewType
    If IsNumeric(value) Then
        OlViewTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olTableView": OlViewTypeFromString = olTableView
        Case "olCardView": OlViewTypeFromString = olCardView
        Case "olCalendarView": OlViewTypeFromString = olCalendarView
        Case "olIconView": OlViewTypeFromString = olIconView
        Case "olTimelineView": OlViewTypeFromString = olTimelineView
        Case "olBusinessCardView": OlViewTypeFromString = olBusinessCardView
        Case "olDailyTaskListView": OlViewTypeFromString = olDailyTaskListView
    End Select
End Function

Function OlViewTypeToString(value As OlViewType) As String
    Select Case value
        Case olTableView: OlViewTypeToString = "olTableView"
        Case olCardView: OlViewTypeToString = "olCardView"
        Case olCalendarView: OlViewTypeToString = "olCalendarView"
        Case olIconView: OlViewTypeToString = "olIconView"
        Case olTimelineView: OlViewTypeToString = "olTimelineView"
        Case olBusinessCardView: OlViewTypeToString = "olBusinessCardView"
        Case olDailyTaskListView: OlViewTypeToString = "olDailyTaskListView"
    End Select
End Function
