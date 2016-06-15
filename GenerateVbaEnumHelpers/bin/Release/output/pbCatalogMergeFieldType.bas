Attribute VB_Name = "wpbCatalogMergeFieldType"
Function pbCatalogMergeFieldTypeFromString(value As String) As pbCatalogMergeFieldType
    If IsNumeric(value) Then
        pbCatalogMergeFieldTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbCatalogMergeFieldTypeText": pbCatalogMergeFieldTypeFromString = pbCatalogMergeFieldTypeText
        Case "pbCatalogMergeFieldTypePicture": pbCatalogMergeFieldTypeFromString = pbCatalogMergeFieldTypePicture
    End Select
End Function

Function pbCatalogMergeFieldTypeToString(value As pbCatalogMergeFieldType) As String
    Select Case value
        Case pbCatalogMergeFieldTypeText: pbCatalogMergeFieldTypeToString = "pbCatalogMergeFieldTypeText"
        Case pbCatalogMergeFieldTypePicture: pbCatalogMergeFieldTypeToString = "pbCatalogMergeFieldTypePicture"
    End Select
End Function
