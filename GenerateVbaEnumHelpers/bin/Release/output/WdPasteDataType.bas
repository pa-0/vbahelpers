Attribute VB_Name = "wWdPasteDataType"
Function WdPasteDataTypeFromString(value As String) As WdPasteDataType
    If IsNumeric(value) Then
        WdPasteDataTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdPasteOLEObject": WdPasteDataTypeFromString = wdPasteOLEObject
        Case "wdPasteRTF": WdPasteDataTypeFromString = wdPasteRTF
        Case "wdPasteText": WdPasteDataTypeFromString = wdPasteText
        Case "wdPasteMetafilePicture": WdPasteDataTypeFromString = wdPasteMetafilePicture
        Case "wdPasteBitmap": WdPasteDataTypeFromString = wdPasteBitmap
        Case "wdPasteDeviceIndependentBitmap": WdPasteDataTypeFromString = wdPasteDeviceIndependentBitmap
        Case "wdPasteHyperlink": WdPasteDataTypeFromString = wdPasteHyperlink
        Case "wdPasteShape": WdPasteDataTypeFromString = wdPasteShape
        Case "wdPasteEnhancedMetafile": WdPasteDataTypeFromString = wdPasteEnhancedMetafile
        Case "wdPasteHTML": WdPasteDataTypeFromString = wdPasteHTML
    End Select
End Function

Function WdPasteDataTypeToString(value As WdPasteDataType) As String
    Select Case value
        Case wdPasteOLEObject: WdPasteDataTypeToString = "wdPasteOLEObject"
        Case wdPasteRTF: WdPasteDataTypeToString = "wdPasteRTF"
        Case wdPasteText: WdPasteDataTypeToString = "wdPasteText"
        Case wdPasteMetafilePicture: WdPasteDataTypeToString = "wdPasteMetafilePicture"
        Case wdPasteBitmap: WdPasteDataTypeToString = "wdPasteBitmap"
        Case wdPasteDeviceIndependentBitmap: WdPasteDataTypeToString = "wdPasteDeviceIndependentBitmap"
        Case wdPasteHyperlink: WdPasteDataTypeToString = "wdPasteHyperlink"
        Case wdPasteShape: WdPasteDataTypeToString = "wdPasteShape"
        Case wdPasteEnhancedMetafile: WdPasteDataTypeToString = "wdPasteEnhancedMetafile"
        Case wdPasteHTML: WdPasteDataTypeToString = "wdPasteHTML"
    End Select
End Function
