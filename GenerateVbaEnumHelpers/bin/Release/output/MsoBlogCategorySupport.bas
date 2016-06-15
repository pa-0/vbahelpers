Attribute VB_Name = "wMsoBlogCategorySupport"
Function MsoBlogCategorySupportFromString(value As String) As MsoBlogCategorySupport
    If IsNumeric(value) Then
        MsoBlogCategorySupportFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoBlogNoCategories": MsoBlogCategorySupportFromString = msoBlogNoCategories
        Case "msoBlogOneCategory": MsoBlogCategorySupportFromString = msoBlogOneCategory
        Case "msoBlogMultipleCategories": MsoBlogCategorySupportFromString = msoBlogMultipleCategories
    End Select
End Function

Function MsoBlogCategorySupportToString(value As MsoBlogCategorySupport) As String
    Select Case value
        Case msoBlogNoCategories: MsoBlogCategorySupportToString = "msoBlogNoCategories"
        Case msoBlogOneCategory: MsoBlogCategorySupportToString = "msoBlogOneCategory"
        Case msoBlogMultipleCategories: MsoBlogCategorySupportToString = "msoBlogMultipleCategories"
    End Select
End Function
