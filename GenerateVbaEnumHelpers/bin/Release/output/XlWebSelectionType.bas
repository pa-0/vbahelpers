Attribute VB_Name = "wXlWebSelectionType"
Function XlWebSelectionTypeFromString(value As String) As XlWebSelectionType
    If IsNumeric(value) Then
        XlWebSelectionTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlEntirePage": XlWebSelectionTypeFromString = xlEntirePage
        Case "xlAllTables": XlWebSelectionTypeFromString = xlAllTables
        Case "xlSpecifiedTables": XlWebSelectionTypeFromString = xlSpecifiedTables
    End Select
End Function

Function XlWebSelectionTypeToString(value As XlWebSelectionType) As String
    Select Case value
        Case xlEntirePage: XlWebSelectionTypeToString = "xlEntirePage"
        Case xlAllTables: XlWebSelectionTypeToString = "xlAllTables"
        Case xlSpecifiedTables: XlWebSelectionTypeToString = "xlSpecifiedTables"
    End Select
End Function
