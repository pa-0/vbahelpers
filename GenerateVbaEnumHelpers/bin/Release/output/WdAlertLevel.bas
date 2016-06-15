Attribute VB_Name = "wWdAlertLevel"
Function WdAlertLevelFromString(value As String) As WdAlertLevel
    If IsNumeric(value) Then
        WdAlertLevelFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdAlertsNone": WdAlertLevelFromString = wdAlertsNone
        Case "wdAlertsMessageBox": WdAlertLevelFromString = wdAlertsMessageBox
        Case "wdAlertsAll": WdAlertLevelFromString = wdAlertsAll
    End Select
End Function

Function WdAlertLevelToString(value As WdAlertLevel) As String
    Select Case value
        Case wdAlertsNone: WdAlertLevelToString = "wdAlertsNone"
        Case wdAlertsMessageBox: WdAlertLevelToString = "wdAlertsMessageBox"
        Case wdAlertsAll: WdAlertLevelToString = "wdAlertsAll"
    End Select
End Function
