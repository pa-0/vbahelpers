Attribute VB_Name = "wMsoCalloutDropType"
Function MsoCalloutDropTypeFromString(value As String) As MsoCalloutDropType
    If IsNumeric(value) Then
        MsoCalloutDropTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoCalloutDropCustom": MsoCalloutDropTypeFromString = msoCalloutDropCustom
        Case "msoCalloutDropTop": MsoCalloutDropTypeFromString = msoCalloutDropTop
        Case "msoCalloutDropCenter": MsoCalloutDropTypeFromString = msoCalloutDropCenter
        Case "msoCalloutDropBottom": MsoCalloutDropTypeFromString = msoCalloutDropBottom
        Case "msoCalloutDropMixed": MsoCalloutDropTypeFromString = msoCalloutDropMixed
    End Select
End Function

Function MsoCalloutDropTypeToString(value As MsoCalloutDropType) As String
    Select Case value
        Case msoCalloutDropCustom: MsoCalloutDropTypeToString = "msoCalloutDropCustom"
        Case msoCalloutDropTop: MsoCalloutDropTypeToString = "msoCalloutDropTop"
        Case msoCalloutDropCenter: MsoCalloutDropTypeToString = "msoCalloutDropCenter"
        Case msoCalloutDropBottom: MsoCalloutDropTypeToString = "msoCalloutDropBottom"
        Case msoCalloutDropMixed: MsoCalloutDropTypeToString = "msoCalloutDropMixed"
    End Select
End Function
