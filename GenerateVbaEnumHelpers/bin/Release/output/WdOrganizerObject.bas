Attribute VB_Name = "wWdOrganizerObject"
Function WdOrganizerObjectFromString(value As String) As WdOrganizerObject
    If IsNumeric(value) Then
        WdOrganizerObjectFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdOrganizerObjectStyles": WdOrganizerObjectFromString = wdOrganizerObjectStyles
        Case "wdOrganizerObjectAutoText": WdOrganizerObjectFromString = wdOrganizerObjectAutoText
        Case "wdOrganizerObjectCommandBars": WdOrganizerObjectFromString = wdOrganizerObjectCommandBars
        Case "wdOrganizerObjectProjectItems": WdOrganizerObjectFromString = wdOrganizerObjectProjectItems
    End Select
End Function

Function WdOrganizerObjectToString(value As WdOrganizerObject) As String
    Select Case value
        Case wdOrganizerObjectStyles: WdOrganizerObjectToString = "wdOrganizerObjectStyles"
        Case wdOrganizerObjectAutoText: WdOrganizerObjectToString = "wdOrganizerObjectAutoText"
        Case wdOrganizerObjectCommandBars: WdOrganizerObjectToString = "wdOrganizerObjectCommandBars"
        Case wdOrganizerObjectProjectItems: WdOrganizerObjectToString = "wdOrganizerObjectProjectItems"
    End Select
End Function
