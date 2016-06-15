Attribute VB_Name = "wMsoParagraphAlignment"
Function MsoParagraphAlignmentFromString(value As String) As MsoParagraphAlignment
    If IsNumeric(value) Then
        MsoParagraphAlignmentFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoAlignLeft": MsoParagraphAlignmentFromString = msoAlignLeft
        Case "msoAlignCenter": MsoParagraphAlignmentFromString = msoAlignCenter
        Case "msoAlignRight": MsoParagraphAlignmentFromString = msoAlignRight
        Case "msoAlignJustify": MsoParagraphAlignmentFromString = msoAlignJustify
        Case "msoAlignDistribute": MsoParagraphAlignmentFromString = msoAlignDistribute
        Case "msoAlignThaiDistribute": MsoParagraphAlignmentFromString = msoAlignThaiDistribute
        Case "msoAlignJustifyLow": MsoParagraphAlignmentFromString = msoAlignJustifyLow
        Case "msoAlignMixed": MsoParagraphAlignmentFromString = msoAlignMixed
    End Select
End Function

Function MsoParagraphAlignmentToString(value As MsoParagraphAlignment) As String
    Select Case value
        Case msoAlignLeft: MsoParagraphAlignmentToString = "msoAlignLeft"
        Case msoAlignCenter: MsoParagraphAlignmentToString = "msoAlignCenter"
        Case msoAlignRight: MsoParagraphAlignmentToString = "msoAlignRight"
        Case msoAlignJustify: MsoParagraphAlignmentToString = "msoAlignJustify"
        Case msoAlignDistribute: MsoParagraphAlignmentToString = "msoAlignDistribute"
        Case msoAlignThaiDistribute: MsoParagraphAlignmentToString = "msoAlignThaiDistribute"
        Case msoAlignJustifyLow: MsoParagraphAlignmentToString = "msoAlignJustifyLow"
        Case msoAlignMixed: MsoParagraphAlignmentToString = "msoAlignMixed"
    End Select
End Function
