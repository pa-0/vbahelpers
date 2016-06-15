Attribute VB_Name = "wWdCustomLabelPageSize"
Function WdCustomLabelPageSizeFromString(value As String) As WdCustomLabelPageSize
    If IsNumeric(value) Then
        WdCustomLabelPageSizeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdCustomLabelLetter": WdCustomLabelPageSizeFromString = wdCustomLabelLetter
        Case "wdCustomLabelLetterLS": WdCustomLabelPageSizeFromString = wdCustomLabelLetterLS
        Case "wdCustomLabelA4": WdCustomLabelPageSizeFromString = wdCustomLabelA4
        Case "wdCustomLabelA4LS": WdCustomLabelPageSizeFromString = wdCustomLabelA4LS
        Case "wdCustomLabelA5": WdCustomLabelPageSizeFromString = wdCustomLabelA5
        Case "wdCustomLabelA5LS": WdCustomLabelPageSizeFromString = wdCustomLabelA5LS
        Case "wdCustomLabelB5": WdCustomLabelPageSizeFromString = wdCustomLabelB5
        Case "wdCustomLabelMini": WdCustomLabelPageSizeFromString = wdCustomLabelMini
        Case "wdCustomLabelFanfold": WdCustomLabelPageSizeFromString = wdCustomLabelFanfold
        Case "wdCustomLabelVertHalfSheet": WdCustomLabelPageSizeFromString = wdCustomLabelVertHalfSheet
        Case "wdCustomLabelVertHalfSheetLS": WdCustomLabelPageSizeFromString = wdCustomLabelVertHalfSheetLS
        Case "wdCustomLabelHigaki": WdCustomLabelPageSizeFromString = wdCustomLabelHigaki
        Case "wdCustomLabelHigakiLS": WdCustomLabelPageSizeFromString = wdCustomLabelHigakiLS
        Case "wdCustomLabelB4JIS": WdCustomLabelPageSizeFromString = wdCustomLabelB4JIS
    End Select
End Function

Function WdCustomLabelPageSizeToString(value As WdCustomLabelPageSize) As String
    Select Case value
        Case wdCustomLabelLetter: WdCustomLabelPageSizeToString = "wdCustomLabelLetter"
        Case wdCustomLabelLetterLS: WdCustomLabelPageSizeToString = "wdCustomLabelLetterLS"
        Case wdCustomLabelA4: WdCustomLabelPageSizeToString = "wdCustomLabelA4"
        Case wdCustomLabelA4LS: WdCustomLabelPageSizeToString = "wdCustomLabelA4LS"
        Case wdCustomLabelA5: WdCustomLabelPageSizeToString = "wdCustomLabelA5"
        Case wdCustomLabelA5LS: WdCustomLabelPageSizeToString = "wdCustomLabelA5LS"
        Case wdCustomLabelB5: WdCustomLabelPageSizeToString = "wdCustomLabelB5"
        Case wdCustomLabelMini: WdCustomLabelPageSizeToString = "wdCustomLabelMini"
        Case wdCustomLabelFanfold: WdCustomLabelPageSizeToString = "wdCustomLabelFanfold"
        Case wdCustomLabelVertHalfSheet: WdCustomLabelPageSizeToString = "wdCustomLabelVertHalfSheet"
        Case wdCustomLabelVertHalfSheetLS: WdCustomLabelPageSizeToString = "wdCustomLabelVertHalfSheetLS"
        Case wdCustomLabelHigaki: WdCustomLabelPageSizeToString = "wdCustomLabelHigaki"
        Case wdCustomLabelHigakiLS: WdCustomLabelPageSizeToString = "wdCustomLabelHigakiLS"
        Case wdCustomLabelB4JIS: WdCustomLabelPageSizeToString = "wdCustomLabelB4JIS"
    End Select
End Function
