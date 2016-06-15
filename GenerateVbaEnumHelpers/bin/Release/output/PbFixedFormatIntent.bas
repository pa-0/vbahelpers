Attribute VB_Name = "wPbFixedFormatIntent"
Function PbFixedFormatIntentFromString(value As String) As PbFixedFormatIntent
    If IsNumeric(value) Then
        PbFixedFormatIntentFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbIntentMinimum": PbFixedFormatIntentFromString = pbIntentMinimum
        Case "pbIntentStandard": PbFixedFormatIntentFromString = pbIntentStandard
        Case "pbIntentPrinting": PbFixedFormatIntentFromString = pbIntentPrinting
        Case "pbIntentCommercial": PbFixedFormatIntentFromString = pbIntentCommercial
    End Select
End Function

Function PbFixedFormatIntentToString(value As PbFixedFormatIntent) As String
    Select Case value
        Case pbIntentMinimum: PbFixedFormatIntentToString = "pbIntentMinimum"
        Case pbIntentStandard: PbFixedFormatIntentToString = "pbIntentStandard"
        Case pbIntentPrinting: PbFixedFormatIntentToString = "pbIntentPrinting"
        Case pbIntentCommercial: PbFixedFormatIntentToString = "pbIntentCommercial"
    End Select
End Function
