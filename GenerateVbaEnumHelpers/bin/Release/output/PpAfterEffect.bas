Attribute VB_Name = "wPpAfterEffect"
Function PpAfterEffectFromString(value As String) As PpAfterEffect
    If IsNumeric(value) Then
        PpAfterEffectFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "ppAfterEffectNothing": PpAfterEffectFromString = ppAfterEffectNothing
        Case "ppAfterEffectHide": PpAfterEffectFromString = ppAfterEffectHide
        Case "ppAfterEffectDim": PpAfterEffectFromString = ppAfterEffectDim
        Case "ppAfterEffectHideOnClick": PpAfterEffectFromString = ppAfterEffectHideOnClick
        Case "ppAfterEffectMixed": PpAfterEffectFromString = ppAfterEffectMixed
    End Select
End Function

Function PpAfterEffectToString(value As PpAfterEffect) As String
    Select Case value
        Case ppAfterEffectNothing: PpAfterEffectToString = "ppAfterEffectNothing"
        Case ppAfterEffectHide: PpAfterEffectToString = "ppAfterEffectHide"
        Case ppAfterEffectDim: PpAfterEffectToString = "ppAfterEffectDim"
        Case ppAfterEffectHideOnClick: PpAfterEffectToString = "ppAfterEffectHideOnClick"
        Case ppAfterEffectMixed: PpAfterEffectToString = "ppAfterEffectMixed"
    End Select
End Function
