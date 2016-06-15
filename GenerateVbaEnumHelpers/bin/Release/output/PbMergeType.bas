Attribute VB_Name = "wPbMergeType"
Function PbMergeTypeFromString(value As String) As PbMergeType
    If IsNumeric(value) Then
        PbMergeTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbMergeDefault": PbMergeTypeFromString = pbMergeDefault
        Case "pbMailMerge": PbMergeTypeFromString = pbMailMerge
        Case "pbCatalogMerge": PbMergeTypeFromString = pbCatalogMerge
        Case "pbEmailMerge": PbMergeTypeFromString = pbEmailMerge
    End Select
End Function

Function PbMergeTypeToString(value As PbMergeType) As String
    Select Case value
        Case pbMergeDefault: PbMergeTypeToString = "pbMergeDefault"
        Case pbMailMerge: PbMergeTypeToString = "pbMailMerge"
        Case pbCatalogMerge: PbMergeTypeToString = "pbCatalogMerge"
        Case pbEmailMerge: PbMergeTypeToString = "pbEmailMerge"
    End Select
End Function
