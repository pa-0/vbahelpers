Attribute VB_Name = "wXlCubeFieldType"
Function XlCubeFieldTypeFromString(value As String) As XlCubeFieldType
    If IsNumeric(value) Then
        XlCubeFieldTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlHierarchy": XlCubeFieldTypeFromString = xlHierarchy
        Case "xlMeasure": XlCubeFieldTypeFromString = xlMeasure
        Case "xlSet": XlCubeFieldTypeFromString = xlSet
    End Select
End Function

Function XlCubeFieldTypeToString(value As XlCubeFieldType) As String
    Select Case value
        Case xlHierarchy: XlCubeFieldTypeToString = "xlHierarchy"
        Case xlMeasure: XlCubeFieldTypeToString = "xlMeasure"
        Case xlSet: XlCubeFieldTypeToString = "xlSet"
    End Select
End Function
