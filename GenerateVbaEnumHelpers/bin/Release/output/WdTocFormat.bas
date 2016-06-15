Attribute VB_Name = "wWdTocFormat"
Function WdTocFormatFromString(value As String) As WdTocFormat
    If IsNumeric(value) Then
        WdTocFormatFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdTOCTemplate": WdTocFormatFromString = wdTOCTemplate
        Case "wdTOCClassic": WdTocFormatFromString = wdTOCClassic
        Case "wdTOCDistinctive": WdTocFormatFromString = wdTOCDistinctive
        Case "wdTOCFancy": WdTocFormatFromString = wdTOCFancy
        Case "wdTOCModern": WdTocFormatFromString = wdTOCModern
        Case "wdTOCFormal": WdTocFormatFromString = wdTOCFormal
        Case "wdTOCSimple": WdTocFormatFromString = wdTOCSimple
    End Select
End Function

Function WdTocFormatToString(value As WdTocFormat) As String
    Select Case value
        Case wdTOCTemplate: WdTocFormatToString = "wdTOCTemplate"
        Case wdTOCClassic: WdTocFormatToString = "wdTOCClassic"
        Case wdTOCDistinctive: WdTocFormatToString = "wdTOCDistinctive"
        Case wdTOCFancy: WdTocFormatToString = "wdTOCFancy"
        Case wdTOCModern: WdTocFormatToString = "wdTOCModern"
        Case wdTOCFormal: WdTocFormatToString = "wdTOCFormal"
        Case wdTOCSimple: WdTocFormatToString = "wdTOCSimple"
    End Select
End Function
