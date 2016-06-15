Attribute VB_Name = "wXlCorruptLoad"
Function XlCorruptLoadFromString(value As String) As XlCorruptLoad
    If IsNumeric(value) Then
        XlCorruptLoadFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlNormalLoad": XlCorruptLoadFromString = xlNormalLoad
        Case "xlRepairFile": XlCorruptLoadFromString = xlRepairFile
        Case "xlExtractData": XlCorruptLoadFromString = xlExtractData
    End Select
End Function

Function XlCorruptLoadToString(value As XlCorruptLoad) As String
    Select Case value
        Case xlNormalLoad: XlCorruptLoadToString = "xlNormalLoad"
        Case xlRepairFile: XlCorruptLoadToString = "xlRepairFile"
        Case xlExtractData: XlCorruptLoadToString = "xlExtractData"
    End Select
End Function
