Attribute VB_Name = "wWdTofFormat"
Function WdTofFormatFromString(value As String) As WdTofFormat
    If IsNumeric(value) Then
        WdTofFormatFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdTOFTemplate": WdTofFormatFromString = wdTOFTemplate
        Case "wdTOFClassic": WdTofFormatFromString = wdTOFClassic
        Case "wdTOFDistinctive": WdTofFormatFromString = wdTOFDistinctive
        Case "wdTOFCentered": WdTofFormatFromString = wdTOFCentered
        Case "wdTOFFormal": WdTofFormatFromString = wdTOFFormal
        Case "wdTOFSimple": WdTofFormatFromString = wdTOFSimple
    End Select
End Function

Function WdTofFormatToString(value As WdTofFormat) As String
    Select Case value
        Case wdTOFTemplate: WdTofFormatToString = "wdTOFTemplate"
        Case wdTOFClassic: WdTofFormatToString = "wdTOFClassic"
        Case wdTOFDistinctive: WdTofFormatToString = "wdTOFDistinctive"
        Case wdTOFCentered: WdTofFormatToString = "wdTOFCentered"
        Case wdTOFFormal: WdTofFormatToString = "wdTOFFormal"
        Case wdTOFSimple: WdTofFormatToString = "wdTOFSimple"
    End Select
End Function
