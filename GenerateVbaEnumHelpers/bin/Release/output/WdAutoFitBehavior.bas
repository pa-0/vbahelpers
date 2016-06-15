Attribute VB_Name = "wWdAutoFitBehavior"
Function WdAutoFitBehaviorFromString(value As String) As WdAutoFitBehavior
    If IsNumeric(value) Then
        WdAutoFitBehaviorFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdAutoFitFixed": WdAutoFitBehaviorFromString = wdAutoFitFixed
        Case "wdAutoFitContent": WdAutoFitBehaviorFromString = wdAutoFitContent
        Case "wdAutoFitWindow": WdAutoFitBehaviorFromString = wdAutoFitWindow
    End Select
End Function

Function WdAutoFitBehaviorToString(value As WdAutoFitBehavior) As String
    Select Case value
        Case wdAutoFitFixed: WdAutoFitBehaviorToString = "wdAutoFitFixed"
        Case wdAutoFitContent: WdAutoFitBehaviorToString = "wdAutoFitContent"
        Case wdAutoFitWindow: WdAutoFitBehaviorToString = "wdAutoFitWindow"
    End Select
End Function
