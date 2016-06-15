Attribute VB_Name = "wXlSmartTagDisplayMode"
Function XlSmartTagDisplayModeFromString(value As String) As XlSmartTagDisplayMode
    If IsNumeric(value) Then
        XlSmartTagDisplayModeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlIndicatorAndButton": XlSmartTagDisplayModeFromString = xlIndicatorAndButton
        Case "xlDisplayNone": XlSmartTagDisplayModeFromString = xlDisplayNone
        Case "xlButtonOnly": XlSmartTagDisplayModeFromString = xlButtonOnly
    End Select
End Function

Function XlSmartTagDisplayModeToString(value As XlSmartTagDisplayMode) As String
    Select Case value
        Case xlIndicatorAndButton: XlSmartTagDisplayModeToString = "xlIndicatorAndButton"
        Case xlDisplayNone: XlSmartTagDisplayModeToString = "xlDisplayNone"
        Case xlButtonOnly: XlSmartTagDisplayModeToString = "xlButtonOnly"
    End Select
End Function
