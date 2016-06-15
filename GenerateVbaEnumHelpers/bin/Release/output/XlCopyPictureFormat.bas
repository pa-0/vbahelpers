Attribute VB_Name = "wXlCopyPictureFormat"
Function XlCopyPictureFormatFromString(value As String) As XlCopyPictureFormat
    If IsNumeric(value) Then
        XlCopyPictureFormatFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlBitmap": XlCopyPictureFormatFromString = xlBitmap
        Case "xlPicture": XlCopyPictureFormatFromString = xlPicture
    End Select
End Function

Function XlCopyPictureFormatToString(value As XlCopyPictureFormat) As String
    Select Case value
        Case xlBitmap: XlCopyPictureFormatToString = "xlBitmap"
        Case xlPicture: XlCopyPictureFormatToString = "xlPicture"
    End Select
End Function
