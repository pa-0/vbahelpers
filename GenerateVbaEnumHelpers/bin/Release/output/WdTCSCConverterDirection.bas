Attribute VB_Name = "wWdTCSCConverterDirection"
Function WdTCSCConverterDirectionFromString(value As String) As WdTCSCConverterDirection
    If IsNumeric(value) Then
        WdTCSCConverterDirectionFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdTCSCConverterDirectionSCTC": WdTCSCConverterDirectionFromString = wdTCSCConverterDirectionSCTC
        Case "wdTCSCConverterDirectionTCSC": WdTCSCConverterDirectionFromString = wdTCSCConverterDirectionTCSC
        Case "wdTCSCConverterDirectionAuto": WdTCSCConverterDirectionFromString = wdTCSCConverterDirectionAuto
    End Select
End Function

Function WdTCSCConverterDirectionToString(value As WdTCSCConverterDirection) As String
    Select Case value
        Case wdTCSCConverterDirectionSCTC: WdTCSCConverterDirectionToString = "wdTCSCConverterDirectionSCTC"
        Case wdTCSCConverterDirectionTCSC: WdTCSCConverterDirectionToString = "wdTCSCConverterDirectionTCSC"
        Case wdTCSCConverterDirectionAuto: WdTCSCConverterDirectionToString = "wdTCSCConverterDirectionAuto"
    End Select
End Function
