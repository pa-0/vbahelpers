Attribute VB_Name = "wMsoIconType"
Function MsoIconTypeFromString(value As String) As MsoIconType
    If IsNumeric(value) Then
        MsoIconTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoIconNone": MsoIconTypeFromString = msoIconNone
        Case "msoIconAlert": MsoIconTypeFromString = msoIconAlert
        Case "msoIconTip": MsoIconTypeFromString = msoIconTip
        Case "msoIconAlertInfo": MsoIconTypeFromString = msoIconAlertInfo
        Case "msoIconAlertWarning": MsoIconTypeFromString = msoIconAlertWarning
        Case "msoIconAlertQuery": MsoIconTypeFromString = msoIconAlertQuery
        Case "msoIconAlertCritical": MsoIconTypeFromString = msoIconAlertCritical
    End Select
End Function

Function MsoIconTypeToString(value As MsoIconType) As String
    Select Case value
        Case msoIconNone: MsoIconTypeToString = "msoIconNone"
        Case msoIconAlert: MsoIconTypeToString = "msoIconAlert"
        Case msoIconTip: MsoIconTypeToString = "msoIconTip"
        Case msoIconAlertInfo: MsoIconTypeToString = "msoIconAlertInfo"
        Case msoIconAlertWarning: MsoIconTypeToString = "msoIconAlertWarning"
        Case msoIconAlertQuery: MsoIconTypeToString = "msoIconAlertQuery"
        Case msoIconAlertCritical: MsoIconTypeToString = "msoIconAlertCritical"
    End Select
End Function
