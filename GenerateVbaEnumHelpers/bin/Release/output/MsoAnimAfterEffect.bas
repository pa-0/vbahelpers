Attribute VB_Name = "wMsoAnimAfterEffect"
Function MsoAnimAfterEffectFromString(value As String) As MsoAnimAfterEffect
    If IsNumeric(value) Then
        MsoAnimAfterEffectFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoAnimAfterEffectNone": MsoAnimAfterEffectFromString = msoAnimAfterEffectNone
        Case "msoAnimAfterEffectDim": MsoAnimAfterEffectFromString = msoAnimAfterEffectDim
        Case "msoAnimAfterEffectHide": MsoAnimAfterEffectFromString = msoAnimAfterEffectHide
        Case "msoAnimAfterEffectHideOnNextClick": MsoAnimAfterEffectFromString = msoAnimAfterEffectHideOnNextClick
        Case "msoAnimAfterEffectMixed": MsoAnimAfterEffectFromString = msoAnimAfterEffectMixed
    End Select
End Function

Function MsoAnimAfterEffectToString(value As MsoAnimAfterEffect) As String
    Select Case value
        Case msoAnimAfterEffectNone: MsoAnimAfterEffectToString = "msoAnimAfterEffectNone"
        Case msoAnimAfterEffectDim: MsoAnimAfterEffectToString = "msoAnimAfterEffectDim"
        Case msoAnimAfterEffectHide: MsoAnimAfterEffectToString = "msoAnimAfterEffectHide"
        Case msoAnimAfterEffectHideOnNextClick: MsoAnimAfterEffectToString = "msoAnimAfterEffectHideOnNextClick"
        Case msoAnimAfterEffectMixed: MsoAnimAfterEffectToString = "msoAnimAfterEffectMixed"
    End Select
End Function
