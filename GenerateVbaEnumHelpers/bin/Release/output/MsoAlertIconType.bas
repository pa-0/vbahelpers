Attribute VB_Name = "wMsoAlertIconType"
Function MsoAlertIconTypeFromString(value As String) As MsoAlertIconType
    If IsNumeric(value) Then
        MsoAlertIconTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoAlertIconNoIcon": MsoAlertIconTypeFromString = msoAlertIconNoIcon
        Case "msoAlertIconCritical": MsoAlertIconTypeFromString = msoAlertIconCritical
        Case "msoAlertIconQuery": MsoAlertIconTypeFromString = msoAlertIconQuery
        Case "msoAlertIconWarning": MsoAlertIconTypeFromString = msoAlertIconWarning
        Case "msoAlertIconInfo": MsoAlertIconTypeFromString = msoAlertIconInfo
    End Select
End Function

Function MsoAlertIconTypeToString(value As MsoAlertIconType) As String
    Select Case value
        Case msoAlertIconNoIcon: MsoAlertIconTypeToString = "msoAlertIconNoIcon"
        Case msoAlertIconCritical: MsoAlertIconTypeToString = "msoAlertIconCritical"
        Case msoAlertIconQuery: MsoAlertIconTypeToString = "msoAlertIconQuery"
        Case msoAlertIconWarning: MsoAlertIconTypeToString = "msoAlertIconWarning"
        Case msoAlertIconInfo: MsoAlertIconTypeToString = "msoAlertIconInfo"
    End Select
End Function
