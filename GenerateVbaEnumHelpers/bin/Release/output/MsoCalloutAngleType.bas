Attribute VB_Name = "wMsoCalloutAngleType"
Function MsoCalloutAngleTypeFromString(value As String) As MsoCalloutAngleType
    If IsNumeric(value) Then
        MsoCalloutAngleTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoCalloutAngleAutomatic": MsoCalloutAngleTypeFromString = msoCalloutAngleAutomatic
        Case "msoCalloutAngle30": MsoCalloutAngleTypeFromString = msoCalloutAngle30
        Case "msoCalloutAngle45": MsoCalloutAngleTypeFromString = msoCalloutAngle45
        Case "msoCalloutAngle60": MsoCalloutAngleTypeFromString = msoCalloutAngle60
        Case "msoCalloutAngle90": MsoCalloutAngleTypeFromString = msoCalloutAngle90
        Case "msoCalloutAngleMixed": MsoCalloutAngleTypeFromString = msoCalloutAngleMixed
    End Select
End Function

Function MsoCalloutAngleTypeToString(value As MsoCalloutAngleType) As String
    Select Case value
        Case msoCalloutAngleAutomatic: MsoCalloutAngleTypeToString = "msoCalloutAngleAutomatic"
        Case msoCalloutAngle30: MsoCalloutAngleTypeToString = "msoCalloutAngle30"
        Case msoCalloutAngle45: MsoCalloutAngleTypeToString = "msoCalloutAngle45"
        Case msoCalloutAngle60: MsoCalloutAngleTypeToString = "msoCalloutAngle60"
        Case msoCalloutAngle90: MsoCalloutAngleTypeToString = "msoCalloutAngle90"
        Case msoCalloutAngleMixed: MsoCalloutAngleTypeToString = "msoCalloutAngleMixed"
    End Select
End Function
