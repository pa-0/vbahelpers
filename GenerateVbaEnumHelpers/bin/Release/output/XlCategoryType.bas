Attribute VB_Name = "wXlCategoryType"
Function XlCategoryTypeFromString(value As String) As XlCategoryType
    If IsNumeric(value) Then
        XlCategoryTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlCategoryScale": XlCategoryTypeFromString = xlCategoryScale
        Case "xlTimeScale": XlCategoryTypeFromString = xlTimeScale
        Case "xlAutomaticScale": XlCategoryTypeFromString = xlAutomaticScale
    End Select
End Function

Function XlCategoryTypeToString(value As XlCategoryType) As String
    Select Case value
        Case xlCategoryScale: XlCategoryTypeToString = "xlCategoryScale"
        Case xlTimeScale: XlCategoryTypeToString = "xlTimeScale"
        Case xlAutomaticScale: XlCategoryTypeToString = "xlAutomaticScale"
    End Select
End Function
