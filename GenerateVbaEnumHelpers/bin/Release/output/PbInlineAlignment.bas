Attribute VB_Name = "wPbInlineAlignment"
Function PbInlineAlignmentFromString(value As String) As PbInlineAlignment
    If IsNumeric(value) Then
        PbInlineAlignmentFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbInlineAlignmentCharacter": PbInlineAlignmentFromString = pbInlineAlignmentCharacter
        Case "pbInlineAlignmentLeft": PbInlineAlignmentFromString = pbInlineAlignmentLeft
        Case "pbInlineAlignmentRight": PbInlineAlignmentFromString = pbInlineAlignmentRight
        Case "pbInlineAlignmentMixed": PbInlineAlignmentFromString = pbInlineAlignmentMixed
    End Select
End Function

Function PbInlineAlignmentToString(value As PbInlineAlignment) As String
    Select Case value
        Case pbInlineAlignmentCharacter: PbInlineAlignmentToString = "pbInlineAlignmentCharacter"
        Case pbInlineAlignmentLeft: PbInlineAlignmentToString = "pbInlineAlignmentLeft"
        Case pbInlineAlignmentRight: PbInlineAlignmentToString = "pbInlineAlignmentRight"
        Case pbInlineAlignmentMixed: PbInlineAlignmentToString = "pbInlineAlignmentMixed"
    End Select
End Function
