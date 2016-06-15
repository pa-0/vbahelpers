Attribute VB_Name = "wOlTimelineViewMode"
Function OlTimelineViewModeFromString(value As String) As OlTimelineViewMode
    If IsNumeric(value) Then
        OlTimelineViewModeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olTimelineViewDay": OlTimelineViewModeFromString = olTimelineViewDay
        Case "olTimelineViewWeek": OlTimelineViewModeFromString = olTimelineViewWeek
        Case "olTimelineViewMonth": OlTimelineViewModeFromString = olTimelineViewMonth
    End Select
End Function

Function OlTimelineViewModeToString(value As OlTimelineViewMode) As String
    Select Case value
        Case olTimelineViewDay: OlTimelineViewModeToString = "olTimelineViewDay"
        Case olTimelineViewWeek: OlTimelineViewModeToString = "olTimelineViewWeek"
        Case olTimelineViewMonth: OlTimelineViewModeToString = "olTimelineViewMonth"
    End Select
End Function
