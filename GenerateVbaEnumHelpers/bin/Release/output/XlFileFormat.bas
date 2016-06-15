Attribute VB_Name = "wXlFileFormat"
Function XlFileFormatFromString(value As String) As XlFileFormat
    If IsNumeric(value) Then
        XlFileFormatFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlSYLK": XlFileFormatFromString = xlSYLK
        Case "xlWKS": XlFileFormatFromString = xlWKS
        Case "xlWK1": XlFileFormatFromString = xlWK1
        Case "xlCSV": XlFileFormatFromString = xlCSV
        Case "xlDBF2": XlFileFormatFromString = xlDBF2
        Case "xlDBF3": XlFileFormatFromString = xlDBF3
        Case "xlDIF": XlFileFormatFromString = xlDIF
        Case "xlDBF4": XlFileFormatFromString = xlDBF4
        Case "xlWJ2WD1": XlFileFormatFromString = xlWJ2WD1
        Case "xlWK3": XlFileFormatFromString = xlWK3
        Case "xlExcel2": XlFileFormatFromString = xlExcel2
        Case "xlTemplate": XlFileFormatFromString = xlTemplate
        Case "xlTemplate8": XlFileFormatFromString = xlTemplate8
        Case "xlAddIn": XlFileFormatFromString = xlAddIn
        Case "xlAddIn8": XlFileFormatFromString = xlAddIn8
        Case "xlTextMac": XlFileFormatFromString = xlTextMac
        Case "xlTextWindows": XlFileFormatFromString = xlTextWindows
        Case "xlTextMSDOS": XlFileFormatFromString = xlTextMSDOS
        Case "xlCSVMac": XlFileFormatFromString = xlCSVMac
        Case "xlCSVWindows": XlFileFormatFromString = xlCSVWindows
        Case "xlCSVMSDOS": XlFileFormatFromString = xlCSVMSDOS
        Case "xlIntlMacro": XlFileFormatFromString = xlIntlMacro
        Case "xlIntlAddIn": XlFileFormatFromString = xlIntlAddIn
        Case "xlExcel2FarEast": XlFileFormatFromString = xlExcel2FarEast
        Case "xlWorks2FarEast": XlFileFormatFromString = xlWorks2FarEast
        Case "xlExcel3": XlFileFormatFromString = xlExcel3
        Case "xlWK1FMT": XlFileFormatFromString = xlWK1FMT
        Case "xlWK1ALL": XlFileFormatFromString = xlWK1ALL
        Case "xlWK3FM3": XlFileFormatFromString = xlWK3FM3
        Case "xlExcel4": XlFileFormatFromString = xlExcel4
        Case "xlWQ1": XlFileFormatFromString = xlWQ1
        Case "xlExcel4Workbook": XlFileFormatFromString = xlExcel4Workbook
        Case "xlTextPrinter": XlFileFormatFromString = xlTextPrinter
        Case "xlWK4": XlFileFormatFromString = xlWK4
        Case "xlExcel7": XlFileFormatFromString = xlExcel7
        Case "xlExcel5": XlFileFormatFromString = xlExcel5
        Case "xlWJ3": XlFileFormatFromString = xlWJ3
        Case "xlWJ3FJ3": XlFileFormatFromString = xlWJ3FJ3
        Case "xlUnicodeText": XlFileFormatFromString = xlUnicodeText
        Case "xlExcel9795": XlFileFormatFromString = xlExcel9795
        Case "xlHtml": XlFileFormatFromString = xlHtml
        Case "xlWebArchive": XlFileFormatFromString = xlWebArchive
        Case "xlXMLSpreadsheet": XlFileFormatFromString = xlXMLSpreadsheet
        Case "xlExcel12": XlFileFormatFromString = xlExcel12
        Case "xlOpenXMLWorkbook": XlFileFormatFromString = xlOpenXMLWorkbook
        Case "xlWorkbookDefault": XlFileFormatFromString = xlWorkbookDefault
        Case "xlOpenXMLWorkbookMacroEnabled": XlFileFormatFromString = xlOpenXMLWorkbookMacroEnabled
        Case "xlOpenXMLTemplateMacroEnabled": XlFileFormatFromString = xlOpenXMLTemplateMacroEnabled
        Case "xlOpenXMLTemplate": XlFileFormatFromString = xlOpenXMLTemplate
        Case "xlOpenXMLAddIn": XlFileFormatFromString = xlOpenXMLAddIn
        Case "xlExcel8": XlFileFormatFromString = xlExcel8
        Case "xlOpenDocumentSpreadsheet": XlFileFormatFromString = xlOpenDocumentSpreadsheet
        Case "xlCurrentPlatformText": XlFileFormatFromString = xlCurrentPlatformText
        Case "xlWorkbookNormal": XlFileFormatFromString = xlWorkbookNormal
    End Select
End Function

Function XlFileFormatToString(value As XlFileFormat) As String
    Select Case value
        Case xlSYLK: XlFileFormatToString = "xlSYLK"
        Case xlWKS: XlFileFormatToString = "xlWKS"
        Case xlWK1: XlFileFormatToString = "xlWK1"
        Case xlCSV: XlFileFormatToString = "xlCSV"
        Case xlDBF2: XlFileFormatToString = "xlDBF2"
        Case xlDBF3: XlFileFormatToString = "xlDBF3"
        Case xlDIF: XlFileFormatToString = "xlDIF"
        Case xlDBF4: XlFileFormatToString = "xlDBF4"
        Case xlWJ2WD1: XlFileFormatToString = "xlWJ2WD1"
        Case xlWK3: XlFileFormatToString = "xlWK3"
        Case xlExcel2: XlFileFormatToString = "xlExcel2"
        Case xlTemplate: XlFileFormatToString = "xlTemplate"
        Case xlTemplate8: XlFileFormatToString = "xlTemplate8"
        Case xlAddIn: XlFileFormatToString = "xlAddIn"
        Case xlAddIn8: XlFileFormatToString = "xlAddIn8"
        Case xlTextMac: XlFileFormatToString = "xlTextMac"
        Case xlTextWindows: XlFileFormatToString = "xlTextWindows"
        Case xlTextMSDOS: XlFileFormatToString = "xlTextMSDOS"
        Case xlCSVMac: XlFileFormatToString = "xlCSVMac"
        Case xlCSVWindows: XlFileFormatToString = "xlCSVWindows"
        Case xlCSVMSDOS: XlFileFormatToString = "xlCSVMSDOS"
        Case xlIntlMacro: XlFileFormatToString = "xlIntlMacro"
        Case xlIntlAddIn: XlFileFormatToString = "xlIntlAddIn"
        Case xlExcel2FarEast: XlFileFormatToString = "xlExcel2FarEast"
        Case xlWorks2FarEast: XlFileFormatToString = "xlWorks2FarEast"
        Case xlExcel3: XlFileFormatToString = "xlExcel3"
        Case xlWK1FMT: XlFileFormatToString = "xlWK1FMT"
        Case xlWK1ALL: XlFileFormatToString = "xlWK1ALL"
        Case xlWK3FM3: XlFileFormatToString = "xlWK3FM3"
        Case xlExcel4: XlFileFormatToString = "xlExcel4"
        Case xlWQ1: XlFileFormatToString = "xlWQ1"
        Case xlExcel4Workbook: XlFileFormatToString = "xlExcel4Workbook"
        Case xlTextPrinter: XlFileFormatToString = "xlTextPrinter"
        Case xlWK4: XlFileFormatToString = "xlWK4"
        Case xlExcel7: XlFileFormatToString = "xlExcel7"
        Case xlExcel5: XlFileFormatToString = "xlExcel5"
        Case xlWJ3: XlFileFormatToString = "xlWJ3"
        Case xlWJ3FJ3: XlFileFormatToString = "xlWJ3FJ3"
        Case xlUnicodeText: XlFileFormatToString = "xlUnicodeText"
        Case xlExcel9795: XlFileFormatToString = "xlExcel9795"
        Case xlHtml: XlFileFormatToString = "xlHtml"
        Case xlWebArchive: XlFileFormatToString = "xlWebArchive"
        Case xlXMLSpreadsheet: XlFileFormatToString = "xlXMLSpreadsheet"
        Case xlExcel12: XlFileFormatToString = "xlExcel12"
        Case xlOpenXMLWorkbook: XlFileFormatToString = "xlOpenXMLWorkbook"
        Case xlWorkbookDefault: XlFileFormatToString = "xlWorkbookDefault"
        Case xlOpenXMLWorkbookMacroEnabled: XlFileFormatToString = "xlOpenXMLWorkbookMacroEnabled"
        Case xlOpenXMLTemplateMacroEnabled: XlFileFormatToString = "xlOpenXMLTemplateMacroEnabled"
        Case xlOpenXMLTemplate: XlFileFormatToString = "xlOpenXMLTemplate"
        Case xlOpenXMLAddIn: XlFileFormatToString = "xlOpenXMLAddIn"
        Case xlExcel8: XlFileFormatToString = "xlExcel8"
        Case xlOpenDocumentSpreadsheet: XlFileFormatToString = "xlOpenDocumentSpreadsheet"
        Case xlCurrentPlatformText: XlFileFormatToString = "xlCurrentPlatformText"
        Case xlWorkbookNormal: XlFileFormatToString = "xlWorkbookNormal"
    End Select
End Function
