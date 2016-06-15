Attribute VB_Name = "wPpFixedFormatIntent"
Function PpFixedFormatIntentFromString(value As String) As PpFixedFormatIntent
    If IsNumeric(value) Then
        PpFixedFormatIntentFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "ppFixedFormatIntentScreen": PpFixedFormatIntentFromString = ppFixedFormatIntentScreen
        Case "ppFixedFormatIntentPrint": PpFixedFormatIntentFromString = ppFixedFormatIntentPrint
    End Select
End Function

Function PpFixedFormatIntentToString(value As PpFixedFormatIntent) As String
    Select Case value
        Case ppFixedFormatIntentScreen: PpFixedFormatIntentToString = "ppFixedFormatIntentScreen"
        Case ppFixedFormatIntentPrint: PpFixedFormatIntentToString = "ppFixedFormatIntentPrint"
    End Select
End Function
