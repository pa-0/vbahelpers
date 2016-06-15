Attribute VB_Name = "wMsoAlertCancelType"
Function MsoAlertCancelTypeFromString(value As String) As MsoAlertCancelType
    If IsNumeric(value) Then
        MsoAlertCancelTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoAlertCancelFirst": MsoAlertCancelTypeFromString = msoAlertCancelFirst
        Case "msoAlertCancelSecond": MsoAlertCancelTypeFromString = msoAlertCancelSecond
        Case "msoAlertCancelThird": MsoAlertCancelTypeFromString = msoAlertCancelThird
        Case "msoAlertCancelFourth": MsoAlertCancelTypeFromString = msoAlertCancelFourth
        Case "msoAlertCancelFifth": MsoAlertCancelTypeFromString = msoAlertCancelFifth
        Case "msoAlertCancelDefault": MsoAlertCancelTypeFromString = msoAlertCancelDefault
    End Select
End Function

Function MsoAlertCancelTypeToString(value As MsoAlertCancelType) As String
    Select Case value
        Case msoAlertCancelFirst: MsoAlertCancelTypeToString = "msoAlertCancelFirst"
        Case msoAlertCancelSecond: MsoAlertCancelTypeToString = "msoAlertCancelSecond"
        Case msoAlertCancelThird: MsoAlertCancelTypeToString = "msoAlertCancelThird"
        Case msoAlertCancelFourth: MsoAlertCancelTypeToString = "msoAlertCancelFourth"
        Case msoAlertCancelFifth: MsoAlertCancelTypeToString = "msoAlertCancelFifth"
        Case msoAlertCancelDefault: MsoAlertCancelTypeToString = "msoAlertCancelDefault"
    End Select
End Function
