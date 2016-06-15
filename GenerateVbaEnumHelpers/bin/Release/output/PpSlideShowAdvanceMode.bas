Attribute VB_Name = "wPpSlideShowAdvanceMode"
Function PpSlideShowAdvanceModeFromString(value As String) As PpSlideShowAdvanceMode
    If IsNumeric(value) Then
        PpSlideShowAdvanceModeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "ppSlideShowManualAdvance": PpSlideShowAdvanceModeFromString = ppSlideShowManualAdvance
        Case "ppSlideShowUseSlideTimings": PpSlideShowAdvanceModeFromString = ppSlideShowUseSlideTimings
        Case "ppSlideShowRehearseNewTimings": PpSlideShowAdvanceModeFromString = ppSlideShowRehearseNewTimings
    End Select
End Function

Function PpSlideShowAdvanceModeToString(value As PpSlideShowAdvanceMode) As String
    Select Case value
        Case ppSlideShowManualAdvance: PpSlideShowAdvanceModeToString = "ppSlideShowManualAdvance"
        Case ppSlideShowUseSlideTimings: PpSlideShowAdvanceModeToString = "ppSlideShowUseSlideTimings"
        Case ppSlideShowRehearseNewTimings: PpSlideShowAdvanceModeToString = "ppSlideShowRehearseNewTimings"
    End Select
End Function
