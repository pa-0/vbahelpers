Attribute VB_Name = "wWdSelectionFlags"
Function WdSelectionFlagsFromString(value As String) As WdSelectionFlags
    If IsNumeric(value) Then
        WdSelectionFlagsFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdSelStartActive": WdSelectionFlagsFromString = wdSelStartActive
        Case "wdSelAtEOL": WdSelectionFlagsFromString = wdSelAtEOL
        Case "wdSelOvertype": WdSelectionFlagsFromString = wdSelOvertype
        Case "wdSelActive": WdSelectionFlagsFromString = wdSelActive
        Case "wdSelReplace": WdSelectionFlagsFromString = wdSelReplace
    End Select
End Function

Function WdSelectionFlagsToString(value As WdSelectionFlags) As String
    Select Case value
        Case wdSelStartActive: WdSelectionFlagsToString = "wdSelStartActive"
        Case wdSelAtEOL: WdSelectionFlagsToString = "wdSelAtEOL"
        Case wdSelOvertype: WdSelectionFlagsToString = "wdSelOvertype"
        Case wdSelActive: WdSelectionFlagsToString = "wdSelActive"
        Case wdSelReplace: WdSelectionFlagsToString = "wdSelReplace"
    End Select
End Function
