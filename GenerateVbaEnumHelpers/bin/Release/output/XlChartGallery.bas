Attribute VB_Name = "wXlChartGallery"
Function XlChartGalleryFromString(value As String) As XlChartGallery
    If IsNumeric(value) Then
        XlChartGalleryFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlBuiltIn": XlChartGalleryFromString = xlBuiltIn
        Case "xlUserDefined": XlChartGalleryFromString = xlUserDefined
        Case "xlAnyGallery": XlChartGalleryFromString = xlAnyGallery
    End Select
End Function

Function XlChartGalleryToString(value As XlChartGallery) As String
    Select Case value
        Case xlBuiltIn: XlChartGalleryToString = "xlBuiltIn"
        Case xlUserDefined: XlChartGalleryToString = "xlUserDefined"
        Case xlAnyGallery: XlChartGalleryToString = "xlAnyGallery"
    End Select
End Function
