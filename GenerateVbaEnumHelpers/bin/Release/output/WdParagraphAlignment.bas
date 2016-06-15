Attribute VB_Name = "wWdParagraphAlignment"
Function WdParagraphAlignmentFromString(value As String) As WdParagraphAlignment
    If IsNumeric(value) Then
        WdParagraphAlignmentFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdAlignParagraphLeft": WdParagraphAlignmentFromString = wdAlignParagraphLeft
        Case "wdAlignParagraphCenter": WdParagraphAlignmentFromString = wdAlignParagraphCenter
        Case "wdAlignParagraphRight": WdParagraphAlignmentFromString = wdAlignParagraphRight
        Case "wdAlignParagraphJustify": WdParagraphAlignmentFromString = wdAlignParagraphJustify
        Case "wdAlignParagraphDistribute": WdParagraphAlignmentFromString = wdAlignParagraphDistribute
        Case "wdAlignParagraphJustifyMed": WdParagraphAlignmentFromString = wdAlignParagraphJustifyMed
        Case "wdAlignParagraphJustifyHi": WdParagraphAlignmentFromString = wdAlignParagraphJustifyHi
        Case "wdAlignParagraphJustifyLow": WdParagraphAlignmentFromString = wdAlignParagraphJustifyLow
        Case "wdAlignParagraphThaiJustify": WdParagraphAlignmentFromString = wdAlignParagraphThaiJustify
    End Select
End Function

Function WdParagraphAlignmentToString(value As WdParagraphAlignment) As String
    Select Case value
        Case wdAlignParagraphLeft: WdParagraphAlignmentToString = "wdAlignParagraphLeft"
        Case wdAlignParagraphCenter: WdParagraphAlignmentToString = "wdAlignParagraphCenter"
        Case wdAlignParagraphRight: WdParagraphAlignmentToString = "wdAlignParagraphRight"
        Case wdAlignParagraphJustify: WdParagraphAlignmentToString = "wdAlignParagraphJustify"
        Case wdAlignParagraphDistribute: WdParagraphAlignmentToString = "wdAlignParagraphDistribute"
        Case wdAlignParagraphJustifyMed: WdParagraphAlignmentToString = "wdAlignParagraphJustifyMed"
        Case wdAlignParagraphJustifyHi: WdParagraphAlignmentToString = "wdAlignParagraphJustifyHi"
        Case wdAlignParagraphJustifyLow: WdParagraphAlignmentToString = "wdAlignParagraphJustifyLow"
        Case wdAlignParagraphThaiJustify: WdParagraphAlignmentToString = "wdAlignParagraphThaiJustify"
    End Select
End Function
