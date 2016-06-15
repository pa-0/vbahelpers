Attribute VB_Name = "wWdPageFit"
Function WdPageFitFromString(value As String) As WdPageFit
    If IsNumeric(value) Then
        WdPageFitFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdPageFitNone": WdPageFitFromString = wdPageFitNone
        Case "wdPageFitFullPage": WdPageFitFromString = wdPageFitFullPage
        Case "wdPageFitBestFit": WdPageFitFromString = wdPageFitBestFit
        Case "wdPageFitTextFit": WdPageFitFromString = wdPageFitTextFit
    End Select
End Function

Function WdPageFitToString(value As WdPageFit) As String
    Select Case value
        Case wdPageFitNone: WdPageFitToString = "wdPageFitNone"
        Case wdPageFitFullPage: WdPageFitToString = "wdPageFitFullPage"
        Case wdPageFitBestFit: WdPageFitToString = "wdPageFitBestFit"
        Case wdPageFitTextFit: WdPageFitToString = "wdPageFitTextFit"
    End Select
End Function
