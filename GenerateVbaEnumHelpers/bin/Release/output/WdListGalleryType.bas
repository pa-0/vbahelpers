Attribute VB_Name = "wWdListGalleryType"
Function WdListGalleryTypeFromString(value As String) As WdListGalleryType
    If IsNumeric(value) Then
        WdListGalleryTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdBulletGallery": WdListGalleryTypeFromString = wdBulletGallery
        Case "wdNumberGallery": WdListGalleryTypeFromString = wdNumberGallery
        Case "wdOutlineNumberGallery": WdListGalleryTypeFromString = wdOutlineNumberGallery
    End Select
End Function

Function WdListGalleryTypeToString(value As WdListGalleryType) As String
    Select Case value
        Case wdBulletGallery: WdListGalleryTypeToString = "wdBulletGallery"
        Case wdNumberGallery: WdListGalleryTypeToString = "wdNumberGallery"
        Case wdOutlineNumberGallery: WdListGalleryTypeToString = "wdOutlineNumberGallery"
    End Select
End Function
