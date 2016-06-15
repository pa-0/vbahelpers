Attribute VB_Name = "wXlXmlLoadOption"
Function XlXmlLoadOptionFromString(value As String) As XlXmlLoadOption
    If IsNumeric(value) Then
        XlXmlLoadOptionFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlXmlLoadPromptUser": XlXmlLoadOptionFromString = xlXmlLoadPromptUser
        Case "xlXmlLoadOpenXml": XlXmlLoadOptionFromString = xlXmlLoadOpenXml
        Case "xlXmlLoadImportToList": XlXmlLoadOptionFromString = xlXmlLoadImportToList
        Case "xlXmlLoadMapXml": XlXmlLoadOptionFromString = xlXmlLoadMapXml
    End Select
End Function

Function XlXmlLoadOptionToString(value As XlXmlLoadOption) As String
    Select Case value
        Case xlXmlLoadPromptUser: XlXmlLoadOptionToString = "xlXmlLoadPromptUser"
        Case xlXmlLoadOpenXml: XlXmlLoadOptionToString = "xlXmlLoadOpenXml"
        Case xlXmlLoadImportToList: XlXmlLoadOptionToString = "xlXmlLoadImportToList"
        Case xlXmlLoadMapXml: XlXmlLoadOptionToString = "xlXmlLoadMapXml"
    End Select
End Function
