Attribute VB_Name = "wPpPasteDataType"
Function PpPasteDataTypeFromString(value As String) As PpPasteDataType
    If IsNumeric(value) Then
        PpPasteDataTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "ppPasteDefault": PpPasteDataTypeFromString = ppPasteDefault
        Case "ppPasteBitmap": PpPasteDataTypeFromString = ppPasteBitmap
        Case "ppPasteEnhancedMetafile": PpPasteDataTypeFromString = ppPasteEnhancedMetafile
        Case "ppPasteMetafilePicture": PpPasteDataTypeFromString = ppPasteMetafilePicture
        Case "ppPasteGIF": PpPasteDataTypeFromString = ppPasteGIF
        Case "ppPasteJPG": PpPasteDataTypeFromString = ppPasteJPG
        Case "ppPastePNG": PpPasteDataTypeFromString = ppPastePNG
        Case "ppPasteText": PpPasteDataTypeFromString = ppPasteText
        Case "ppPasteHTML": PpPasteDataTypeFromString = ppPasteHTML
        Case "ppPasteRTF": PpPasteDataTypeFromString = ppPasteRTF
        Case "ppPasteOLEObject": PpPasteDataTypeFromString = ppPasteOLEObject
        Case "ppPasteShape": PpPasteDataTypeFromString = ppPasteShape
    End Select
End Function

Function PpPasteDataTypeToString(value As PpPasteDataType) As String
    Select Case value
        Case ppPasteDefault: PpPasteDataTypeToString = "ppPasteDefault"
        Case ppPasteBitmap: PpPasteDataTypeToString = "ppPasteBitmap"
        Case ppPasteEnhancedMetafile: PpPasteDataTypeToString = "ppPasteEnhancedMetafile"
        Case ppPasteMetafilePicture: PpPasteDataTypeToString = "ppPasteMetafilePicture"
        Case ppPasteGIF: PpPasteDataTypeToString = "ppPasteGIF"
        Case ppPasteJPG: PpPasteDataTypeToString = "ppPasteJPG"
        Case ppPastePNG: PpPasteDataTypeToString = "ppPastePNG"
        Case ppPasteText: PpPasteDataTypeToString = "ppPasteText"
        Case ppPasteHTML: PpPasteDataTypeToString = "ppPasteHTML"
        Case ppPasteRTF: PpPasteDataTypeToString = "ppPasteRTF"
        Case ppPasteOLEObject: PpPasteDataTypeToString = "ppPasteOLEObject"
        Case ppPasteShape: PpPasteDataTypeToString = "ppPasteShape"
    End Select
End Function
