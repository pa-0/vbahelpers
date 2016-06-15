Attribute VB_Name = "wXlLocationInTable"
Function XlLocationInTableFromString(value As String) As XlLocationInTable
    If IsNumeric(value) Then
        XlLocationInTableFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlPageHeader": XlLocationInTableFromString = xlPageHeader
        Case "xlDataHeader": XlLocationInTableFromString = xlDataHeader
        Case "xlRowItem": XlLocationInTableFromString = xlRowItem
        Case "xlColumnItem": XlLocationInTableFromString = xlColumnItem
        Case "xlPageItem": XlLocationInTableFromString = xlPageItem
        Case "xlDataItem": XlLocationInTableFromString = xlDataItem
        Case "xlTableBody": XlLocationInTableFromString = xlTableBody
        Case "xlRowHeader": XlLocationInTableFromString = xlRowHeader
        Case "xlColumnHeader": XlLocationInTableFromString = xlColumnHeader
    End Select
End Function

Function XlLocationInTableToString(value As XlLocationInTable) As String
    Select Case value
        Case xlPageHeader: XlLocationInTableToString = "xlPageHeader"
        Case xlDataHeader: XlLocationInTableToString = "xlDataHeader"
        Case xlRowItem: XlLocationInTableToString = "xlRowItem"
        Case xlColumnItem: XlLocationInTableToString = "xlColumnItem"
        Case xlPageItem: XlLocationInTableToString = "xlPageItem"
        Case xlDataItem: XlLocationInTableToString = "xlDataItem"
        Case xlTableBody: XlLocationInTableToString = "xlTableBody"
        Case xlRowHeader: XlLocationInTableToString = "xlRowHeader"
        Case xlColumnHeader: XlLocationInTableToString = "xlColumnHeader"
    End Select
End Function
