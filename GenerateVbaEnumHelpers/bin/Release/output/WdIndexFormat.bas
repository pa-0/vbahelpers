Attribute VB_Name = "wWdIndexFormat"
Function WdIndexFormatFromString(value As String) As WdIndexFormat
    If IsNumeric(value) Then
        WdIndexFormatFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdIndexTemplate": WdIndexFormatFromString = wdIndexTemplate
        Case "wdIndexClassic": WdIndexFormatFromString = wdIndexClassic
        Case "wdIndexFancy": WdIndexFormatFromString = wdIndexFancy
        Case "wdIndexModern": WdIndexFormatFromString = wdIndexModern
        Case "wdIndexBulleted": WdIndexFormatFromString = wdIndexBulleted
        Case "wdIndexFormal": WdIndexFormatFromString = wdIndexFormal
        Case "wdIndexSimple": WdIndexFormatFromString = wdIndexSimple
    End Select
End Function

Function WdIndexFormatToString(value As WdIndexFormat) As String
    Select Case value
        Case wdIndexTemplate: WdIndexFormatToString = "wdIndexTemplate"
        Case wdIndexClassic: WdIndexFormatToString = "wdIndexClassic"
        Case wdIndexFancy: WdIndexFormatToString = "wdIndexFancy"
        Case wdIndexModern: WdIndexFormatToString = "wdIndexModern"
        Case wdIndexBulleted: WdIndexFormatToString = "wdIndexBulleted"
        Case wdIndexFormal: WdIndexFormatToString = "wdIndexFormal"
        Case wdIndexSimple: WdIndexFormatToString = "wdIndexSimple"
    End Select
End Function
