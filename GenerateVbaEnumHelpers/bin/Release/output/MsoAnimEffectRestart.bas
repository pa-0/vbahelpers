Attribute VB_Name = "wMsoAnimEffectRestart"
Function MsoAnimEffectRestartFromString(value As String) As MsoAnimEffectRestart
    If IsNumeric(value) Then
        MsoAnimEffectRestartFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoAnimEffectRestartAlways": MsoAnimEffectRestartFromString = msoAnimEffectRestartAlways
        Case "msoAnimEffectRestartWhenOff": MsoAnimEffectRestartFromString = msoAnimEffectRestartWhenOff
        Case "msoAnimEffectRestartNever": MsoAnimEffectRestartFromString = msoAnimEffectRestartNever
    End Select
End Function

Function MsoAnimEffectRestartToString(value As MsoAnimEffectRestart) As String
    Select Case value
        Case msoAnimEffectRestartAlways: MsoAnimEffectRestartToString = "msoAnimEffectRestartAlways"
        Case msoAnimEffectRestartWhenOff: MsoAnimEffectRestartToString = "msoAnimEffectRestartWhenOff"
        Case msoAnimEffectRestartNever: MsoAnimEffectRestartToString = "msoAnimEffectRestartNever"
    End Select
End Function
