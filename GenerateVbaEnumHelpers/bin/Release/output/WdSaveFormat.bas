Attribute VB_Name = "wWdSaveFormat"
Function WdSaveFormatFromString(value As String) As WdSaveFormat
    If IsNumeric(value) Then
        WdSaveFormatFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdFormatDocument": WdSaveFormatFromString = wdFormatDocument
        Case "wdFormatDocument97": WdSaveFormatFromString = wdFormatDocument97
        Case "wdFormatTemplate": WdSaveFormatFromString = wdFormatTemplate
        Case "wdFormatTemplate97": WdSaveFormatFromString = wdFormatTemplate97
        Case "wdFormatText": WdSaveFormatFromString = wdFormatText
        Case "wdFormatTextLineBreaks": WdSaveFormatFromString = wdFormatTextLineBreaks
        Case "wdFormatDOSText": WdSaveFormatFromString = wdFormatDOSText
        Case "wdFormatDOSTextLineBreaks": WdSaveFormatFromString = wdFormatDOSTextLineBreaks
        Case "wdFormatRTF": WdSaveFormatFromString = wdFormatRTF
        Case "wdFormatUnicodeText": WdSaveFormatFromString = wdFormatUnicodeText
        Case "wdFormatEncodedText": WdSaveFormatFromString = wdFormatEncodedText
        Case "wdFormatHTML": WdSaveFormatFromString = wdFormatHTML
        Case "wdFormatWebArchive": WdSaveFormatFromString = wdFormatWebArchive
        Case "wdFormatFilteredHTML": WdSaveFormatFromString = wdFormatFilteredHTML
        Case "wdFormatXML": WdSaveFormatFromString = wdFormatXML
        Case "wdFormatXMLDocument": WdSaveFormatFromString = wdFormatXMLDocument
        Case "wdFormatXMLDocumentMacroEnabled": WdSaveFormatFromString = wdFormatXMLDocumentMacroEnabled
        Case "wdFormatXMLTemplate": WdSaveFormatFromString = wdFormatXMLTemplate
        Case "wdFormatXMLTemplateMacroEnabled": WdSaveFormatFromString = wdFormatXMLTemplateMacroEnabled
        Case "wdFormatDocumentDefault": WdSaveFormatFromString = wdFormatDocumentDefault
        Case "wdFormatPDF": WdSaveFormatFromString = wdFormatPDF
        Case "wdFormatXPS": WdSaveFormatFromString = wdFormatXPS
        Case "wdFormatFlatXML": WdSaveFormatFromString = wdFormatFlatXML
        Case "wdFormatFlatXMLMacroEnabled": WdSaveFormatFromString = wdFormatFlatXMLMacroEnabled
        Case "wdFormatFlatXMLTemplate": WdSaveFormatFromString = wdFormatFlatXMLTemplate
        Case "wdFormatFlatXMLTemplateMacroEnabled": WdSaveFormatFromString = wdFormatFlatXMLTemplateMacroEnabled
        Case "wdFormatOpenDocumentText": WdSaveFormatFromString = wdFormatOpenDocumentText
    End Select
End Function

Function WdSaveFormatToString(value As WdSaveFormat) As String
    Select Case value
        Case wdFormatDocument: WdSaveFormatToString = "wdFormatDocument"
        Case wdFormatDocument97: WdSaveFormatToString = "wdFormatDocument97"
        Case wdFormatTemplate: WdSaveFormatToString = "wdFormatTemplate"
        Case wdFormatTemplate97: WdSaveFormatToString = "wdFormatTemplate97"
        Case wdFormatText: WdSaveFormatToString = "wdFormatText"
        Case wdFormatTextLineBreaks: WdSaveFormatToString = "wdFormatTextLineBreaks"
        Case wdFormatDOSText: WdSaveFormatToString = "wdFormatDOSText"
        Case wdFormatDOSTextLineBreaks: WdSaveFormatToString = "wdFormatDOSTextLineBreaks"
        Case wdFormatRTF: WdSaveFormatToString = "wdFormatRTF"
        Case wdFormatUnicodeText: WdSaveFormatToString = "wdFormatUnicodeText"
        Case wdFormatEncodedText: WdSaveFormatToString = "wdFormatEncodedText"
        Case wdFormatHTML: WdSaveFormatToString = "wdFormatHTML"
        Case wdFormatWebArchive: WdSaveFormatToString = "wdFormatWebArchive"
        Case wdFormatFilteredHTML: WdSaveFormatToString = "wdFormatFilteredHTML"
        Case wdFormatXML: WdSaveFormatToString = "wdFormatXML"
        Case wdFormatXMLDocument: WdSaveFormatToString = "wdFormatXMLDocument"
        Case wdFormatXMLDocumentMacroEnabled: WdSaveFormatToString = "wdFormatXMLDocumentMacroEnabled"
        Case wdFormatXMLTemplate: WdSaveFormatToString = "wdFormatXMLTemplate"
        Case wdFormatXMLTemplateMacroEnabled: WdSaveFormatToString = "wdFormatXMLTemplateMacroEnabled"
        Case wdFormatDocumentDefault: WdSaveFormatToString = "wdFormatDocumentDefault"
        Case wdFormatPDF: WdSaveFormatToString = "wdFormatPDF"
        Case wdFormatXPS: WdSaveFormatToString = "wdFormatXPS"
        Case wdFormatFlatXML: WdSaveFormatToString = "wdFormatFlatXML"
        Case wdFormatFlatXMLMacroEnabled: WdSaveFormatToString = "wdFormatFlatXMLMacroEnabled"
        Case wdFormatFlatXMLTemplate: WdSaveFormatToString = "wdFormatFlatXMLTemplate"
        Case wdFormatFlatXMLTemplateMacroEnabled: WdSaveFormatToString = "wdFormatFlatXMLTemplateMacroEnabled"
        Case wdFormatOpenDocumentText: WdSaveFormatToString = "wdFormatOpenDocumentText"
    End Select
End Function
