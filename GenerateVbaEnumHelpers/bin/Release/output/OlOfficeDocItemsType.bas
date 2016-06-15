Attribute VB_Name = "wOlOfficeDocItemsType"
Function OlOfficeDocItemsTypeFromString(value As String) As OlOfficeDocItemsType
    If IsNumeric(value) Then
        OlOfficeDocItemsTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olExcelWorkSheetItem": OlOfficeDocItemsTypeFromString = olExcelWorkSheetItem
        Case "olWordDocumentItem": OlOfficeDocItemsTypeFromString = olWordDocumentItem
        Case "olPowerPointShowItem": OlOfficeDocItemsTypeFromString = olPowerPointShowItem
    End Select
End Function

Function OlOfficeDocItemsTypeToString(value As OlOfficeDocItemsType) As String
    Select Case value
        Case olExcelWorkSheetItem: OlOfficeDocItemsTypeToString = "olExcelWorkSheetItem"
        Case olWordDocumentItem: OlOfficeDocItemsTypeToString = "olWordDocumentItem"
        Case olPowerPointShowItem: OlOfficeDocItemsTypeToString = "olPowerPointShowItem"
    End Select
End Function
