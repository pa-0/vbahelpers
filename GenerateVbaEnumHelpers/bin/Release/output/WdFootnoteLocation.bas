Attribute VB_Name = "wWdFootnoteLocation"
Function WdFootnoteLocationFromString(value As String) As WdFootnoteLocation
    If IsNumeric(value) Then
        WdFootnoteLocationFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdBottomOfPage": WdFootnoteLocationFromString = wdBottomOfPage
        Case "wdBeneathText": WdFootnoteLocationFromString = wdBeneathText
    End Select
End Function

Function WdFootnoteLocationToString(value As WdFootnoteLocation) As String
    Select Case value
        Case wdBottomOfPage: WdFootnoteLocationToString = "wdBottomOfPage"
        Case wdBeneathText: WdFootnoteLocationToString = "wdBeneathText"
    End Select
End Function
