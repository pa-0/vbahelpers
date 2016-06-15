Attribute VB_Name = "wWdToaFormat"
Function WdToaFormatFromString(value As String) As WdToaFormat
    If IsNumeric(value) Then
        WdToaFormatFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdTOATemplate": WdToaFormatFromString = wdTOATemplate
        Case "wdTOAClassic": WdToaFormatFromString = wdTOAClassic
        Case "wdTOADistinctive": WdToaFormatFromString = wdTOADistinctive
        Case "wdTOAFormal": WdToaFormatFromString = wdTOAFormal
        Case "wdTOASimple": WdToaFormatFromString = wdTOASimple
    End Select
End Function

Function WdToaFormatToString(value As WdToaFormat) As String
    Select Case value
        Case wdTOATemplate: WdToaFormatToString = "wdTOATemplate"
        Case wdTOAClassic: WdToaFormatToString = "wdTOAClassic"
        Case wdTOADistinctive: WdToaFormatToString = "wdTOADistinctive"
        Case wdTOAFormal: WdToaFormatToString = "wdTOAFormal"
        Case wdTOASimple: WdToaFormatToString = "wdTOASimple"
    End Select
End Function
