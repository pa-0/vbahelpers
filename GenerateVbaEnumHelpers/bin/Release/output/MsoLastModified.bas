Attribute VB_Name = "wMsoLastModified"
Function MsoLastModifiedFromString(value As String) As MsoLastModified
    If IsNumeric(value) Then
        MsoLastModifiedFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoLastModifiedYesterday": MsoLastModifiedFromString = msoLastModifiedYesterday
        Case "msoLastModifiedToday": MsoLastModifiedFromString = msoLastModifiedToday
        Case "msoLastModifiedLastWeek": MsoLastModifiedFromString = msoLastModifiedLastWeek
        Case "msoLastModifiedThisWeek": MsoLastModifiedFromString = msoLastModifiedThisWeek
        Case "msoLastModifiedLastMonth": MsoLastModifiedFromString = msoLastModifiedLastMonth
        Case "msoLastModifiedThisMonth": MsoLastModifiedFromString = msoLastModifiedThisMonth
        Case "msoLastModifiedAnyTime": MsoLastModifiedFromString = msoLastModifiedAnyTime
    End Select
End Function

Function MsoLastModifiedToString(value As MsoLastModified) As String
    Select Case value
        Case msoLastModifiedYesterday: MsoLastModifiedToString = "msoLastModifiedYesterday"
        Case msoLastModifiedToday: MsoLastModifiedToString = "msoLastModifiedToday"
        Case msoLastModifiedLastWeek: MsoLastModifiedToString = "msoLastModifiedLastWeek"
        Case msoLastModifiedThisWeek: MsoLastModifiedToString = "msoLastModifiedThisWeek"
        Case msoLastModifiedLastMonth: MsoLastModifiedToString = "msoLastModifiedLastMonth"
        Case msoLastModifiedThisMonth: MsoLastModifiedToString = "msoLastModifiedThisMonth"
        Case msoLastModifiedAnyTime: MsoLastModifiedToString = "msoLastModifiedAnyTime"
    End Select
End Function
