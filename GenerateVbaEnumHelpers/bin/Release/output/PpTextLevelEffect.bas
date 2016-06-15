Attribute VB_Name = "wPpTextLevelEffect"
Function PpTextLevelEffectFromString(value As String) As PpTextLevelEffect
    If IsNumeric(value) Then
        PpTextLevelEffectFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "ppAnimateLevelNone": PpTextLevelEffectFromString = ppAnimateLevelNone
        Case "ppAnimateByFirstLevel": PpTextLevelEffectFromString = ppAnimateByFirstLevel
        Case "ppAnimateBySecondLevel": PpTextLevelEffectFromString = ppAnimateBySecondLevel
        Case "ppAnimateByThirdLevel": PpTextLevelEffectFromString = ppAnimateByThirdLevel
        Case "ppAnimateByFourthLevel": PpTextLevelEffectFromString = ppAnimateByFourthLevel
        Case "ppAnimateByFifthLevel": PpTextLevelEffectFromString = ppAnimateByFifthLevel
        Case "ppAnimateByAllLevels": PpTextLevelEffectFromString = ppAnimateByAllLevels
        Case "ppAnimateLevelMixed": PpTextLevelEffectFromString = ppAnimateLevelMixed
    End Select
End Function

Function PpTextLevelEffectToString(value As PpTextLevelEffect) As String
    Select Case value
        Case ppAnimateLevelNone: PpTextLevelEffectToString = "ppAnimateLevelNone"
        Case ppAnimateByFirstLevel: PpTextLevelEffectToString = "ppAnimateByFirstLevel"
        Case ppAnimateBySecondLevel: PpTextLevelEffectToString = "ppAnimateBySecondLevel"
        Case ppAnimateByThirdLevel: PpTextLevelEffectToString = "ppAnimateByThirdLevel"
        Case ppAnimateByFourthLevel: PpTextLevelEffectToString = "ppAnimateByFourthLevel"
        Case ppAnimateByFifthLevel: PpTextLevelEffectToString = "ppAnimateByFifthLevel"
        Case ppAnimateByAllLevels: PpTextLevelEffectToString = "ppAnimateByAllLevels"
        Case ppAnimateLevelMixed: PpTextLevelEffectToString = "ppAnimateLevelMixed"
    End Select
End Function
