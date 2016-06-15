Attribute VB_Name = "wPbMailMergeDataFieldType"
Function PbMailMergeDataFieldTypeFromString(value As String) As PbMailMergeDataFieldType
    If IsNumeric(value) Then
        PbMailMergeDataFieldTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbMailMergeDataFieldString": PbMailMergeDataFieldTypeFromString = pbMailMergeDataFieldString
        Case "pbMailMergeDataFieldPicture": PbMailMergeDataFieldTypeFromString = pbMailMergeDataFieldPicture
    End Select
End Function

Function PbMailMergeDataFieldTypeToString(value As PbMailMergeDataFieldType) As String
    Select Case value
        Case pbMailMergeDataFieldString: PbMailMergeDataFieldTypeToString = "pbMailMergeDataFieldString"
        Case pbMailMergeDataFieldPicture: PbMailMergeDataFieldTypeToString = "pbMailMergeDataFieldPicture"
    End Select
End Function
