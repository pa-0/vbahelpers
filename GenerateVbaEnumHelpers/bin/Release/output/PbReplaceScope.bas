Attribute VB_Name = "wPbReplaceScope"
Function PbReplaceScopeFromString(value As String) As PbReplaceScope
    If IsNumeric(value) Then
        PbReplaceScopeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbReplaceScopeNone": PbReplaceScopeFromString = pbReplaceScopeNone
        Case "pbReplaceScopeOne": PbReplaceScopeFromString = pbReplaceScopeOne
        Case "pbReplaceScopeAll": PbReplaceScopeFromString = pbReplaceScopeAll
    End Select
End Function

Function PbReplaceScopeToString(value As PbReplaceScope) As String
    Select Case value
        Case pbReplaceScopeNone: PbReplaceScopeToString = "pbReplaceScopeNone"
        Case pbReplaceScopeOne: PbReplaceScopeToString = "pbReplaceScopeOne"
        Case pbReplaceScopeAll: PbReplaceScopeToString = "pbReplaceScopeAll"
    End Select
End Function
