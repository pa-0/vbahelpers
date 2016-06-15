Attribute VB_Name = "wXlReferenceType"
Function XlReferenceTypeFromString(value As String) As XlReferenceType
    If IsNumeric(value) Then
        XlReferenceTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlAbsolute": XlReferenceTypeFromString = xlAbsolute
        Case "xlAbsRowRelColumn": XlReferenceTypeFromString = xlAbsRowRelColumn
        Case "xlRelRowAbsColumn": XlReferenceTypeFromString = xlRelRowAbsColumn
        Case "xlRelative": XlReferenceTypeFromString = xlRelative
    End Select
End Function

Function XlReferenceTypeToString(value As XlReferenceType) As String
    Select Case value
        Case xlAbsolute: XlReferenceTypeToString = "xlAbsolute"
        Case xlAbsRowRelColumn: XlReferenceTypeToString = "xlAbsRowRelColumn"
        Case xlRelRowAbsColumn: XlReferenceTypeToString = "xlRelRowAbsColumn"
        Case xlRelative: XlReferenceTypeToString = "xlRelative"
    End Select
End Function
