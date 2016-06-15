Attribute VB_Name = "wMsoTextEffectAlignment"
Function MsoTextEffectAlignmentFromString(value As String) As MsoTextEffectAlignment
    If IsNumeric(value) Then
        MsoTextEffectAlignmentFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoTextEffectAlignmentLeft": MsoTextEffectAlignmentFromString = msoTextEffectAlignmentLeft
        Case "msoTextEffectAlignmentCentered": MsoTextEffectAlignmentFromString = msoTextEffectAlignmentCentered
        Case "msoTextEffectAlignmentRight": MsoTextEffectAlignmentFromString = msoTextEffectAlignmentRight
        Case "msoTextEffectAlignmentLetterJustify": MsoTextEffectAlignmentFromString = msoTextEffectAlignmentLetterJustify
        Case "msoTextEffectAlignmentWordJustify": MsoTextEffectAlignmentFromString = msoTextEffectAlignmentWordJustify
        Case "msoTextEffectAlignmentStretchJustify": MsoTextEffectAlignmentFromString = msoTextEffectAlignmentStretchJustify
        Case "msoTextEffectAlignmentMixed": MsoTextEffectAlignmentFromString = msoTextEffectAlignmentMixed
    End Select
End Function

Function MsoTextEffectAlignmentToString(value As MsoTextEffectAlignment) As String
    Select Case value
        Case msoTextEffectAlignmentLeft: MsoTextEffectAlignmentToString = "msoTextEffectAlignmentLeft"
        Case msoTextEffectAlignmentCentered: MsoTextEffectAlignmentToString = "msoTextEffectAlignmentCentered"
        Case msoTextEffectAlignmentRight: MsoTextEffectAlignmentToString = "msoTextEffectAlignmentRight"
        Case msoTextEffectAlignmentLetterJustify: MsoTextEffectAlignmentToString = "msoTextEffectAlignmentLetterJustify"
        Case msoTextEffectAlignmentWordJustify: MsoTextEffectAlignmentToString = "msoTextEffectAlignmentWordJustify"
        Case msoTextEffectAlignmentStretchJustify: MsoTextEffectAlignmentToString = "msoTextEffectAlignmentStretchJustify"
        Case msoTextEffectAlignmentMixed: MsoTextEffectAlignmentToString = "msoTextEffectAlignmentMixed"
    End Select
End Function
