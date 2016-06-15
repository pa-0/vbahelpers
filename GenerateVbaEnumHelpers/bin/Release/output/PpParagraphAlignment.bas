Attribute VB_Name = "wPpParagraphAlignment"
Function PpParagraphAlignmentFromString(value As String) As PpParagraphAlignment
    If IsNumeric(value) Then
        PpParagraphAlignmentFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "ppAlignLeft": PpParagraphAlignmentFromString = ppAlignLeft
        Case "ppAlignCenter": PpParagraphAlignmentFromString = ppAlignCenter
        Case "ppAlignRight": PpParagraphAlignmentFromString = ppAlignRight
        Case "ppAlignJustify": PpParagraphAlignmentFromString = ppAlignJustify
        Case "ppAlignDistribute": PpParagraphAlignmentFromString = ppAlignDistribute
        Case "ppAlignThaiDistribute": PpParagraphAlignmentFromString = ppAlignThaiDistribute
        Case "ppAlignJustifyLow": PpParagraphAlignmentFromString = ppAlignJustifyLow
        Case "ppAlignmentMixed": PpParagraphAlignmentFromString = ppAlignmentMixed
    End Select
End Function

Function PpParagraphAlignmentToString(value As PpParagraphAlignment) As String
    Select Case value
        Case ppAlignLeft: PpParagraphAlignmentToString = "ppAlignLeft"
        Case ppAlignCenter: PpParagraphAlignmentToString = "ppAlignCenter"
        Case ppAlignRight: PpParagraphAlignmentToString = "ppAlignRight"
        Case ppAlignJustify: PpParagraphAlignmentToString = "ppAlignJustify"
        Case ppAlignDistribute: PpParagraphAlignmentToString = "ppAlignDistribute"
        Case ppAlignThaiDistribute: PpParagraphAlignmentToString = "ppAlignThaiDistribute"
        Case ppAlignJustifyLow: PpParagraphAlignmentToString = "ppAlignJustifyLow"
        Case ppAlignmentMixed: PpParagraphAlignmentToString = "ppAlignmentMixed"
    End Select
End Function
