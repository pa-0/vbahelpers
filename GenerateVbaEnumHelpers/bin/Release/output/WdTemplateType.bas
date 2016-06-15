Attribute VB_Name = "wWdTemplateType"
Function WdTemplateTypeFromString(value As String) As WdTemplateType
    If IsNumeric(value) Then
        WdTemplateTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdNormalTemplate": WdTemplateTypeFromString = wdNormalTemplate
        Case "wdGlobalTemplate": WdTemplateTypeFromString = wdGlobalTemplate
        Case "wdAttachedTemplate": WdTemplateTypeFromString = wdAttachedTemplate
    End Select
End Function

Function WdTemplateTypeToString(value As WdTemplateType) As String
    Select Case value
        Case wdNormalTemplate: WdTemplateTypeToString = "wdNormalTemplate"
        Case wdGlobalTemplate: WdTemplateTypeToString = "wdGlobalTemplate"
        Case wdAttachedTemplate: WdTemplateTypeToString = "wdAttachedTemplate"
    End Select
End Function
