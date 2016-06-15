Attribute VB_Name = "wXlSaveAsAccessMode"
Function XlSaveAsAccessModeFromString(value As String) As XlSaveAsAccessMode
    If IsNumeric(value) Then
        XlSaveAsAccessModeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlNoChange": XlSaveAsAccessModeFromString = xlNoChange
        Case "xlShared": XlSaveAsAccessModeFromString = xlShared
        Case "xlExclusive": XlSaveAsAccessModeFromString = xlExclusive
    End Select
End Function

Function XlSaveAsAccessModeToString(value As XlSaveAsAccessMode) As String
    Select Case value
        Case xlNoChange: XlSaveAsAccessModeToString = "xlNoChange"
        Case xlShared: XlSaveAsAccessModeToString = "xlShared"
        Case xlExclusive: XlSaveAsAccessModeToString = "xlExclusive"
    End Select
End Function
