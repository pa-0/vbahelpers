Attribute VB_Name = "wPpTextUnitEffect"
Function PpTextUnitEffectFromString(value As String) As PpTextUnitEffect
    If IsNumeric(value) Then
        PpTextUnitEffectFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "ppAnimateByParagraph": PpTextUnitEffectFromString = ppAnimateByParagraph
        Case "ppAnimateByWord": PpTextUnitEffectFromString = ppAnimateByWord
        Case "ppAnimateByCharacter": PpTextUnitEffectFromString = ppAnimateByCharacter
        Case "ppAnimateUnitMixed": PpTextUnitEffectFromString = ppAnimateUnitMixed
    End Select
End Function

Function PpTextUnitEffectToString(value As PpTextUnitEffect) As String
    Select Case value
        Case ppAnimateByParagraph: PpTextUnitEffectToString = "ppAnimateByParagraph"
        Case ppAnimateByWord: PpTextUnitEffectToString = "ppAnimateByWord"
        Case ppAnimateByCharacter: PpTextUnitEffectToString = "ppAnimateByCharacter"
        Case ppAnimateUnitMixed: PpTextUnitEffectToString = "ppAnimateUnitMixed"
    End Select
End Function
