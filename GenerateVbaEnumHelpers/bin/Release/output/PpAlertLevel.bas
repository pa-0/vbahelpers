Attribute VB_Name = "wPpAlertLevel"
Function PpAlertLevelFromString(value As String) As PpAlertLevel
    If IsNumeric(value) Then
        PpAlertLevelFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "ppAlertsNone": PpAlertLevelFromString = ppAlertsNone
        Case "ppAlertsAll": PpAlertLevelFromString = ppAlertsAll
    End Select
End Function

Function PpAlertLevelToString(value As PpAlertLevel) As String
    Select Case value
        Case ppAlertsNone: PpAlertLevelToString = "ppAlertsNone"
        Case ppAlertsAll: PpAlertLevelToString = "ppAlertsAll"
    End Select
End Function
