Attribute VB_Name = "wPpAdvanceMode"
Function PpAdvanceModeFromString(value As String) As PpAdvanceMode
    If IsNumeric(value) Then
        PpAdvanceModeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "ppAdvanceOnClick": PpAdvanceModeFromString = ppAdvanceOnClick
        Case "ppAdvanceOnTime": PpAdvanceModeFromString = ppAdvanceOnTime
        Case "ppAdvanceModeMixed": PpAdvanceModeFromString = ppAdvanceModeMixed
    End Select
End Function

Function PpAdvanceModeToString(value As PpAdvanceMode) As String
    Select Case value
        Case ppAdvanceOnClick: PpAdvanceModeToString = "ppAdvanceOnClick"
        Case ppAdvanceOnTime: PpAdvanceModeToString = "ppAdvanceOnTime"
        Case ppAdvanceModeMixed: PpAdvanceModeToString = "ppAdvanceModeMixed"
    End Select
End Function
