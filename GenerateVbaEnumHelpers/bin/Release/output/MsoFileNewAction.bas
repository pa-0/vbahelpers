Attribute VB_Name = "wMsoFileNewAction"
Function MsoFileNewActionFromString(value As String) As MsoFileNewAction
    If IsNumeric(value) Then
        MsoFileNewActionFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoEditFile": MsoFileNewActionFromString = msoEditFile
        Case "msoCreateNewFile": MsoFileNewActionFromString = msoCreateNewFile
        Case "msoOpenFile": MsoFileNewActionFromString = msoOpenFile
    End Select
End Function

Function MsoFileNewActionToString(value As MsoFileNewAction) As String
    Select Case value
        Case msoEditFile: MsoFileNewActionToString = "msoEditFile"
        Case msoCreateNewFile: MsoFileNewActionToString = "msoCreateNewFile"
        Case msoOpenFile: MsoFileNewActionToString = "msoOpenFile"
    End Select
End Function
