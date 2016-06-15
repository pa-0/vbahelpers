Attribute VB_Name = "wMsoAnimEffectAfter"
Function MsoAnimEffectAfterFromString(value As String) As MsoAnimEffectAfter
    If IsNumeric(value) Then
        MsoAnimEffectAfterFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoAnimEffectAfterFreeze": MsoAnimEffectAfterFromString = msoAnimEffectAfterFreeze
        Case "msoAnimEffectAfterRemove": MsoAnimEffectAfterFromString = msoAnimEffectAfterRemove
        Case "msoAnimEffectAfterHold": MsoAnimEffectAfterFromString = msoAnimEffectAfterHold
        Case "msoAnimEffectAfterTransition": MsoAnimEffectAfterFromString = msoAnimEffectAfterTransition
    End Select
End Function

Function MsoAnimEffectAfterToString(value As MsoAnimEffectAfter) As String
    Select Case value
        Case msoAnimEffectAfterFreeze: MsoAnimEffectAfterToString = "msoAnimEffectAfterFreeze"
        Case msoAnimEffectAfterRemove: MsoAnimEffectAfterToString = "msoAnimEffectAfterRemove"
        Case msoAnimEffectAfterHold: MsoAnimEffectAfterToString = "msoAnimEffectAfterHold"
        Case msoAnimEffectAfterTransition: MsoAnimEffectAfterToString = "msoAnimEffectAfterTransition"
    End Select
End Function
