Attribute VB_Name = "wXlHAlign"
Function XlHAlignFromString(value As String) As XlHAlign
    If IsNumeric(value) Then
        XlHAlignFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlHAlignGeneral": XlHAlignFromString = xlHAlignGeneral
        Case "xlHAlignFill": XlHAlignFromString = xlHAlignFill
        Case "xlHAlignCenterAcrossSelection": XlHAlignFromString = xlHAlignCenterAcrossSelection
        Case "xlHAlignRight": XlHAlignFromString = xlHAlignRight
        Case "xlHAlignLeft": XlHAlignFromString = xlHAlignLeft
        Case "xlHAlignJustify": XlHAlignFromString = xlHAlignJustify
        Case "xlHAlignDistributed": XlHAlignFromString = xlHAlignDistributed
        Case "xlHAlignCenter": XlHAlignFromString = xlHAlignCenter
    End Select
End Function

Function XlHAlignToString(value As XlHAlign) As String
    Select Case value
        Case xlHAlignGeneral: XlHAlignToString = "xlHAlignGeneral"
        Case xlHAlignFill: XlHAlignToString = "xlHAlignFill"
        Case xlHAlignCenterAcrossSelection: XlHAlignToString = "xlHAlignCenterAcrossSelection"
        Case xlHAlignRight: XlHAlignToString = "xlHAlignRight"
        Case xlHAlignLeft: XlHAlignToString = "xlHAlignLeft"
        Case xlHAlignJustify: XlHAlignToString = "xlHAlignJustify"
        Case xlHAlignDistributed: XlHAlignToString = "xlHAlignDistributed"
        Case xlHAlignCenter: XlHAlignToString = "xlHAlignCenter"
    End Select
End Function
