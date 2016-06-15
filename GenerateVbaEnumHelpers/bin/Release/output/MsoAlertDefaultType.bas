Attribute VB_Name = "wMsoAlertDefaultType"
Function MsoAlertDefaultTypeFromString(value As String) As MsoAlertDefaultType
    If IsNumeric(value) Then
        MsoAlertDefaultTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoAlertDefaultFirst": MsoAlertDefaultTypeFromString = msoAlertDefaultFirst
        Case "msoAlertDefaultSecond": MsoAlertDefaultTypeFromString = msoAlertDefaultSecond
        Case "msoAlertDefaultThird": MsoAlertDefaultTypeFromString = msoAlertDefaultThird
        Case "msoAlertDefaultFourth": MsoAlertDefaultTypeFromString = msoAlertDefaultFourth
        Case "msoAlertDefaultFifth": MsoAlertDefaultTypeFromString = msoAlertDefaultFifth
    End Select
End Function

Function MsoAlertDefaultTypeToString(value As MsoAlertDefaultType) As String
    Select Case value
        Case msoAlertDefaultFirst: MsoAlertDefaultTypeToString = "msoAlertDefaultFirst"
        Case msoAlertDefaultSecond: MsoAlertDefaultTypeToString = "msoAlertDefaultSecond"
        Case msoAlertDefaultThird: MsoAlertDefaultTypeToString = "msoAlertDefaultThird"
        Case msoAlertDefaultFourth: MsoAlertDefaultTypeToString = "msoAlertDefaultFourth"
        Case msoAlertDefaultFifth: MsoAlertDefaultTypeToString = "msoAlertDefaultFifth"
    End Select
End Function
