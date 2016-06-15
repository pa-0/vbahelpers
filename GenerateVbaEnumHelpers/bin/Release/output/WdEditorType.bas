Attribute VB_Name = "wWdEditorType"
Function WdEditorTypeFromString(value As String) As WdEditorType
    If IsNumeric(value) Then
        WdEditorTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdEditorCurrent": WdEditorTypeFromString = wdEditorCurrent
        Case "wdEditorEditors": WdEditorTypeFromString = wdEditorEditors
        Case "wdEditorOwners": WdEditorTypeFromString = wdEditorOwners
        Case "wdEditorEveryone": WdEditorTypeFromString = wdEditorEveryone
    End Select
End Function

Function WdEditorTypeToString(value As WdEditorType) As String
    Select Case value
        Case wdEditorCurrent: WdEditorTypeToString = "wdEditorCurrent"
        Case wdEditorEditors: WdEditorTypeToString = "wdEditorEditors"
        Case wdEditorOwners: WdEditorTypeToString = "wdEditorOwners"
        Case wdEditorEveryone: WdEditorTypeToString = "wdEditorEveryone"
    End Select
End Function
