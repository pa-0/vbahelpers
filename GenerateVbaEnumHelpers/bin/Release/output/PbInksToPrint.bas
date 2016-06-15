Attribute VB_Name = "wPbInksToPrint"
Function PbInksToPrintFromString(value As String) As PbInksToPrint
    If IsNumeric(value) Then
        PbInksToPrintFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbInksToPrintAll": PbInksToPrintFromString = pbInksToPrintAll
        Case "pbInksToPrintUsed": PbInksToPrintFromString = pbInksToPrintUsed
        Case "pbInksToPrintConvertSpotToProcess": PbInksToPrintFromString = pbInksToPrintConvertSpotToProcess
    End Select
End Function

Function PbInksToPrintToString(value As PbInksToPrint) As String
    Select Case value
        Case pbInksToPrintAll: PbInksToPrintToString = "pbInksToPrintAll"
        Case pbInksToPrintUsed: PbInksToPrintToString = "pbInksToPrintUsed"
        Case pbInksToPrintConvertSpotToProcess: PbInksToPrintToString = "pbInksToPrintConvertSpotToProcess"
    End Select
End Function
