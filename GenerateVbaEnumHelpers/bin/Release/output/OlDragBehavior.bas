Attribute VB_Name = "wOlDragBehavior"
Function OlDragBehaviorFromString(value As String) As OlDragBehavior
    If IsNumeric(value) Then
        OlDragBehaviorFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olDragBehaviorDisabled": OlDragBehaviorFromString = olDragBehaviorDisabled
        Case "olDragBehaviorEnabled": OlDragBehaviorFromString = olDragBehaviorEnabled
    End Select
End Function

Function OlDragBehaviorToString(value As OlDragBehavior) As String
    Select Case value
        Case olDragBehaviorDisabled: OlDragBehaviorToString = "olDragBehaviorDisabled"
        Case olDragBehaviorEnabled: OlDragBehaviorToString = "olDragBehaviorEnabled"
    End Select
End Function
