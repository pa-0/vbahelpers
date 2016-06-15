Attribute VB_Name = "wWdUpdateStyleListBehavior"
Function WdUpdateStyleListBehaviorFromString(value As String) As WdUpdateStyleListBehavior
    If IsNumeric(value) Then
        WdUpdateStyleListBehaviorFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdListBehaviorKeepPreviousPattern": WdUpdateStyleListBehaviorFromString = wdListBehaviorKeepPreviousPattern
        Case "wdListBehaviorAddBulletsNumbering": WdUpdateStyleListBehaviorFromString = wdListBehaviorAddBulletsNumbering
    End Select
End Function

Function WdUpdateStyleListBehaviorToString(value As WdUpdateStyleListBehavior) As String
    Select Case value
        Case wdListBehaviorKeepPreviousPattern: WdUpdateStyleListBehaviorToString = "wdListBehaviorKeepPreviousPattern"
        Case wdListBehaviorAddBulletsNumbering: WdUpdateStyleListBehaviorToString = "wdListBehaviorAddBulletsNumbering"
    End Select
End Function
