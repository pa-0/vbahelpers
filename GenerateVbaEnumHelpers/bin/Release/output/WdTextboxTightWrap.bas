Attribute VB_Name = "wWdTextboxTightWrap"
Function WdTextboxTightWrapFromString(value As String) As WdTextboxTightWrap
    If IsNumeric(value) Then
        WdTextboxTightWrapFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdTightNone": WdTextboxTightWrapFromString = wdTightNone
        Case "wdTightAll": WdTextboxTightWrapFromString = wdTightAll
        Case "wdTightFirstAndLastLines": WdTextboxTightWrapFromString = wdTightFirstAndLastLines
        Case "wdTightFirstLineOnly": WdTextboxTightWrapFromString = wdTightFirstLineOnly
        Case "wdTightLastLineOnly": WdTextboxTightWrapFromString = wdTightLastLineOnly
    End Select
End Function

Function WdTextboxTightWrapToString(value As WdTextboxTightWrap) As String
    Select Case value
        Case wdTightNone: WdTextboxTightWrapToString = "wdTightNone"
        Case wdTightAll: WdTextboxTightWrapToString = "wdTightAll"
        Case wdTightFirstAndLastLines: WdTextboxTightWrapToString = "wdTightFirstAndLastLines"
        Case wdTightFirstLineOnly: WdTextboxTightWrapToString = "wdTightFirstLineOnly"
        Case wdTightLastLineOnly: WdTextboxTightWrapToString = "wdTightLastLineOnly"
    End Select
End Function
