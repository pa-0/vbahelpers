Attribute VB_Name = "wMsoPresetLightingDirection"
Function MsoPresetLightingDirectionFromString(value As String) As MsoPresetLightingDirection
    If IsNumeric(value) Then
        MsoPresetLightingDirectionFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoLightingTopLeft": MsoPresetLightingDirectionFromString = msoLightingTopLeft
        Case "msoLightingTop": MsoPresetLightingDirectionFromString = msoLightingTop
        Case "msoLightingTopRight": MsoPresetLightingDirectionFromString = msoLightingTopRight
        Case "msoLightingLeft": MsoPresetLightingDirectionFromString = msoLightingLeft
        Case "msoLightingNone": MsoPresetLightingDirectionFromString = msoLightingNone
        Case "msoLightingRight": MsoPresetLightingDirectionFromString = msoLightingRight
        Case "msoLightingBottomLeft": MsoPresetLightingDirectionFromString = msoLightingBottomLeft
        Case "msoLightingBottom": MsoPresetLightingDirectionFromString = msoLightingBottom
        Case "msoLightingBottomRight": MsoPresetLightingDirectionFromString = msoLightingBottomRight
        Case "msoPresetLightingDirectionMixed": MsoPresetLightingDirectionFromString = msoPresetLightingDirectionMixed
    End Select
End Function

Function MsoPresetLightingDirectionToString(value As MsoPresetLightingDirection) As String
    Select Case value
        Case msoLightingTopLeft: MsoPresetLightingDirectionToString = "msoLightingTopLeft"
        Case msoLightingTop: MsoPresetLightingDirectionToString = "msoLightingTop"
        Case msoLightingTopRight: MsoPresetLightingDirectionToString = "msoLightingTopRight"
        Case msoLightingLeft: MsoPresetLightingDirectionToString = "msoLightingLeft"
        Case msoLightingNone: MsoPresetLightingDirectionToString = "msoLightingNone"
        Case msoLightingRight: MsoPresetLightingDirectionToString = "msoLightingRight"
        Case msoLightingBottomLeft: MsoPresetLightingDirectionToString = "msoLightingBottomLeft"
        Case msoLightingBottom: MsoPresetLightingDirectionToString = "msoLightingBottom"
        Case msoLightingBottomRight: MsoPresetLightingDirectionToString = "msoLightingBottomRight"
        Case msoPresetLightingDirectionMixed: MsoPresetLightingDirectionToString = "msoPresetLightingDirectionMixed"
    End Select
End Function
