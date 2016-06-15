Attribute VB_Name = "wWdEndnoteLocation"
Function WdEndnoteLocationFromString(value As String) As WdEndnoteLocation
    If IsNumeric(value) Then
        WdEndnoteLocationFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdEndOfSection": WdEndnoteLocationFromString = wdEndOfSection
        Case "wdEndOfDocument": WdEndnoteLocationFromString = wdEndOfDocument
    End Select
End Function

Function WdEndnoteLocationToString(value As WdEndnoteLocation) As String
    Select Case value
        Case wdEndOfSection: WdEndnoteLocationToString = "wdEndOfSection"
        Case wdEndOfDocument: WdEndnoteLocationToString = "wdEndOfDocument"
    End Select
End Function
