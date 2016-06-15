Attribute VB_Name = "wXlWebFormatting"
Function XlWebFormattingFromString(value As String) As XlWebFormatting
    If IsNumeric(value) Then
        XlWebFormattingFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlWebFormattingAll": XlWebFormattingFromString = xlWebFormattingAll
        Case "xlWebFormattingRTF": XlWebFormattingFromString = xlWebFormattingRTF
        Case "xlWebFormattingNone": XlWebFormattingFromString = xlWebFormattingNone
    End Select
End Function

Function XlWebFormattingToString(value As XlWebFormatting) As String
    Select Case value
        Case xlWebFormattingAll: XlWebFormattingToString = "xlWebFormattingAll"
        Case xlWebFormattingRTF: XlWebFormattingToString = "xlWebFormattingRTF"
        Case xlWebFormattingNone: XlWebFormattingToString = "xlWebFormattingNone"
    End Select
End Function
