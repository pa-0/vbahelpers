Attribute VB_Name = "wMsoPresetExtrusionDirection"
Function MsoPresetExtrusionDirectionFromString(value As String) As MsoPresetExtrusionDirection
    If IsNumeric(value) Then
        MsoPresetExtrusionDirectionFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoExtrusionBottomRight": MsoPresetExtrusionDirectionFromString = msoExtrusionBottomRight
        Case "msoExtrusionBottom": MsoPresetExtrusionDirectionFromString = msoExtrusionBottom
        Case "msoExtrusionBottomLeft": MsoPresetExtrusionDirectionFromString = msoExtrusionBottomLeft
        Case "msoExtrusionRight": MsoPresetExtrusionDirectionFromString = msoExtrusionRight
        Case "msoExtrusionNone": MsoPresetExtrusionDirectionFromString = msoExtrusionNone
        Case "msoExtrusionLeft": MsoPresetExtrusionDirectionFromString = msoExtrusionLeft
        Case "msoExtrusionTopRight": MsoPresetExtrusionDirectionFromString = msoExtrusionTopRight
        Case "msoExtrusionTop": MsoPresetExtrusionDirectionFromString = msoExtrusionTop
        Case "msoExtrusionTopLeft": MsoPresetExtrusionDirectionFromString = msoExtrusionTopLeft
        Case "msoPresetExtrusionDirectionMixed": MsoPresetExtrusionDirectionFromString = msoPresetExtrusionDirectionMixed
    End Select
End Function

Function MsoPresetExtrusionDirectionToString(value As MsoPresetExtrusionDirection) As String
    Select Case value
        Case msoExtrusionBottomRight: MsoPresetExtrusionDirectionToString = "msoExtrusionBottomRight"
        Case msoExtrusionBottom: MsoPresetExtrusionDirectionToString = "msoExtrusionBottom"
        Case msoExtrusionBottomLeft: MsoPresetExtrusionDirectionToString = "msoExtrusionBottomLeft"
        Case msoExtrusionRight: MsoPresetExtrusionDirectionToString = "msoExtrusionRight"
        Case msoExtrusionNone: MsoPresetExtrusionDirectionToString = "msoExtrusionNone"
        Case msoExtrusionLeft: MsoPresetExtrusionDirectionToString = "msoExtrusionLeft"
        Case msoExtrusionTopRight: MsoPresetExtrusionDirectionToString = "msoExtrusionTopRight"
        Case msoExtrusionTop: MsoPresetExtrusionDirectionToString = "msoExtrusionTop"
        Case msoExtrusionTopLeft: MsoPresetExtrusionDirectionToString = "msoExtrusionTopLeft"
        Case msoPresetExtrusionDirectionMixed: MsoPresetExtrusionDirectionToString = "msoPresetExtrusionDirectionMixed"
    End Select
End Function
