Attribute VB_Name = "wWdLinkType"
Function WdLinkTypeFromString(value As String) As WdLinkType
    If IsNumeric(value) Then
        WdLinkTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdLinkTypeOLE": WdLinkTypeFromString = wdLinkTypeOLE
        Case "wdLinkTypePicture": WdLinkTypeFromString = wdLinkTypePicture
        Case "wdLinkTypeText": WdLinkTypeFromString = wdLinkTypeText
        Case "wdLinkTypeReference": WdLinkTypeFromString = wdLinkTypeReference
        Case "wdLinkTypeInclude": WdLinkTypeFromString = wdLinkTypeInclude
        Case "wdLinkTypeImport": WdLinkTypeFromString = wdLinkTypeImport
        Case "wdLinkTypeDDE": WdLinkTypeFromString = wdLinkTypeDDE
        Case "wdLinkTypeDDEAuto": WdLinkTypeFromString = wdLinkTypeDDEAuto
        Case "wdLinkTypeChart": WdLinkTypeFromString = wdLinkTypeChart
    End Select
End Function

Function WdLinkTypeToString(value As WdLinkType) As String
    Select Case value
        Case wdLinkTypeOLE: WdLinkTypeToString = "wdLinkTypeOLE"
        Case wdLinkTypePicture: WdLinkTypeToString = "wdLinkTypePicture"
        Case wdLinkTypeText: WdLinkTypeToString = "wdLinkTypeText"
        Case wdLinkTypeReference: WdLinkTypeToString = "wdLinkTypeReference"
        Case wdLinkTypeInclude: WdLinkTypeToString = "wdLinkTypeInclude"
        Case wdLinkTypeImport: WdLinkTypeToString = "wdLinkTypeImport"
        Case wdLinkTypeDDE: WdLinkTypeToString = "wdLinkTypeDDE"
        Case wdLinkTypeDDEAuto: WdLinkTypeToString = "wdLinkTypeDDEAuto"
        Case wdLinkTypeChart: WdLinkTypeToString = "wdLinkTypeChart"
    End Select
End Function
