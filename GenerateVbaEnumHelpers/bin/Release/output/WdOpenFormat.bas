Attribute VB_Name = "wWdOpenFormat"
Function WdOpenFormatFromString(value As String) As WdOpenFormat
    If IsNumeric(value) Then
        WdOpenFormatFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdOpenFormatAuto": WdOpenFormatFromString = wdOpenFormatAuto
        Case "wdOpenFormatDocument": WdOpenFormatFromString = wdOpenFormatDocument
        Case "wdOpenFormatDocument97": WdOpenFormatFromString = wdOpenFormatDocument97
        Case "wdOpenFormatTemplate": WdOpenFormatFromString = wdOpenFormatTemplate
        Case "wdOpenFormatTemplate97": WdOpenFormatFromString = wdOpenFormatTemplate97
        Case "wdOpenFormatRTF": WdOpenFormatFromString = wdOpenFormatRTF
        Case "wdOpenFormatText": WdOpenFormatFromString = wdOpenFormatText
        Case "wdOpenFormatUnicodeText": WdOpenFormatFromString = wdOpenFormatUnicodeText
        Case "wdOpenFormatEncodedText": WdOpenFormatFromString = wdOpenFormatEncodedText
        Case "wdOpenFormatAllWord": WdOpenFormatFromString = wdOpenFormatAllWord
        Case "wdOpenFormatWebPages": WdOpenFormatFromString = wdOpenFormatWebPages
        Case "wdOpenFormatXML": WdOpenFormatFromString = wdOpenFormatXML
        Case "wdOpenFormatXMLDocument": WdOpenFormatFromString = wdOpenFormatXMLDocument
        Case "wdOpenFormatXMLDocumentMacroEnabled": WdOpenFormatFromString = wdOpenFormatXMLDocumentMacroEnabled
        Case "wdOpenFormatXMLTemplate": WdOpenFormatFromString = wdOpenFormatXMLTemplate
        Case "wdOpenFormatXMLTemplateMacroEnabled": WdOpenFormatFromString = wdOpenFormatXMLTemplateMacroEnabled
        Case "wdOpenFormatAllWordTemplates": WdOpenFormatFromString = wdOpenFormatAllWordTemplates
        Case "wdOpenFormatXMLDocumentSerialized": WdOpenFormatFromString = wdOpenFormatXMLDocumentSerialized
        Case "wdOpenFormatXMLDocumentMacroEnabledSerialized": WdOpenFormatFromString = wdOpenFormatXMLDocumentMacroEnabledSerialized
        Case "wdOpenFormatXMLTemplateSerialized": WdOpenFormatFromString = wdOpenFormatXMLTemplateSerialized
        Case "wdOpenFormatXMLTemplateMacroEnabledSerialized": WdOpenFormatFromString = wdOpenFormatXMLTemplateMacroEnabledSerialized
        Case "wdOpenFormatOpenDocumentText": WdOpenFormatFromString = wdOpenFormatOpenDocumentText
    End Select
End Function

Function WdOpenFormatToString(value As WdOpenFormat) As String
    Select Case value
        Case wdOpenFormatAuto: WdOpenFormatToString = "wdOpenFormatAuto"
        Case wdOpenFormatDocument: WdOpenFormatToString = "wdOpenFormatDocument"
        Case wdOpenFormatDocument97: WdOpenFormatToString = "wdOpenFormatDocument97"
        Case wdOpenFormatTemplate: WdOpenFormatToString = "wdOpenFormatTemplate"
        Case wdOpenFormatTemplate97: WdOpenFormatToString = "wdOpenFormatTemplate97"
        Case wdOpenFormatRTF: WdOpenFormatToString = "wdOpenFormatRTF"
        Case wdOpenFormatText: WdOpenFormatToString = "wdOpenFormatText"
        Case wdOpenFormatUnicodeText: WdOpenFormatToString = "wdOpenFormatUnicodeText"
        Case wdOpenFormatEncodedText: WdOpenFormatToString = "wdOpenFormatEncodedText"
        Case wdOpenFormatAllWord: WdOpenFormatToString = "wdOpenFormatAllWord"
        Case wdOpenFormatWebPages: WdOpenFormatToString = "wdOpenFormatWebPages"
        Case wdOpenFormatXML: WdOpenFormatToString = "wdOpenFormatXML"
        Case wdOpenFormatXMLDocument: WdOpenFormatToString = "wdOpenFormatXMLDocument"
        Case wdOpenFormatXMLDocumentMacroEnabled: WdOpenFormatToString = "wdOpenFormatXMLDocumentMacroEnabled"
        Case wdOpenFormatXMLTemplate: WdOpenFormatToString = "wdOpenFormatXMLTemplate"
        Case wdOpenFormatXMLTemplateMacroEnabled: WdOpenFormatToString = "wdOpenFormatXMLTemplateMacroEnabled"
        Case wdOpenFormatAllWordTemplates: WdOpenFormatToString = "wdOpenFormatAllWordTemplates"
        Case wdOpenFormatXMLDocumentSerialized: WdOpenFormatToString = "wdOpenFormatXMLDocumentSerialized"
        Case wdOpenFormatXMLDocumentMacroEnabledSerialized: WdOpenFormatToString = "wdOpenFormatXMLDocumentMacroEnabledSerialized"
        Case wdOpenFormatXMLTemplateSerialized: WdOpenFormatToString = "wdOpenFormatXMLTemplateSerialized"
        Case wdOpenFormatXMLTemplateMacroEnabledSerialized: WdOpenFormatToString = "wdOpenFormatXMLTemplateMacroEnabledSerialized"
        Case wdOpenFormatOpenDocumentText: WdOpenFormatToString = "wdOpenFormatOpenDocumentText"
    End Select
End Function
