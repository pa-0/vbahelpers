Attribute VB_Name = "wXlSaveAction"
Function XlSaveActionFromString(value As String) As XlSaveAction
    If IsNumeric(value) Then
        XlSaveActionFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlSaveChanges": XlSaveActionFromString = xlSaveChanges
        Case "xlDoNotSaveChanges": XlSaveActionFromString = xlDoNotSaveChanges
    End Select
End Function

Function XlSaveActionToString(value As XlSaveAction) As String
    Select Case value
        Case xlSaveChanges: XlSaveActionToString = "xlSaveChanges"
        Case xlDoNotSaveChanges: XlSaveActionToString = "xlDoNotSaveChanges"
    End Select
End Function
