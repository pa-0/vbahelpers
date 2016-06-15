Attribute VB_Name = "wOlFormatYesNo"
Function OlFormatYesNoFromString(value As String) As OlFormatYesNo
    If IsNumeric(value) Then
        OlFormatYesNoFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olFormatYesNoYesNo": OlFormatYesNoFromString = olFormatYesNoYesNo
        Case "olFormatYesNoOnOff": OlFormatYesNoFromString = olFormatYesNoOnOff
        Case "olFormatYesNoTrueFalse": OlFormatYesNoFromString = olFormatYesNoTrueFalse
        Case "olFormatYesNoIcon": OlFormatYesNoFromString = olFormatYesNoIcon
    End Select
End Function

Function OlFormatYesNoToString(value As OlFormatYesNo) As String
    Select Case value
        Case olFormatYesNoYesNo: OlFormatYesNoToString = "olFormatYesNoYesNo"
        Case olFormatYesNoOnOff: OlFormatYesNoToString = "olFormatYesNoOnOff"
        Case olFormatYesNoTrueFalse: OlFormatYesNoToString = "olFormatYesNoTrueFalse"
        Case olFormatYesNoIcon: OlFormatYesNoToString = "olFormatYesNoIcon"
    End Select
End Function
