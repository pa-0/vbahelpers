Attribute VB_Name = "wWdViewTypeOld"
Function WdViewTypeOldFromString(value As String) As WdViewTypeOld
    If IsNumeric(value) Then
        WdViewTypeOldFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdPageView": WdViewTypeOldFromString = wdPageView
        Case "wdOnlineView": WdViewTypeOldFromString = wdOnlineView
    End Select
End Function

Function WdViewTypeOldToString(value As WdViewTypeOld) As String
    Select Case value
        Case wdPageView: WdViewTypeOldToString = "wdPageView"
        Case wdOnlineView: WdViewTypeOldToString = "wdOnlineView"
    End Select
End Function
