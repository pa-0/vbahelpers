Attribute VB_Name = "wWdInlineShapeType"
Function WdInlineShapeTypeFromString(value As String) As WdInlineShapeType
    If IsNumeric(value) Then
        WdInlineShapeTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdInlineShapeEmbeddedOLEObject": WdInlineShapeTypeFromString = wdInlineShapeEmbeddedOLEObject
        Case "wdInlineShapeLinkedOLEObject": WdInlineShapeTypeFromString = wdInlineShapeLinkedOLEObject
        Case "wdInlineShapePicture": WdInlineShapeTypeFromString = wdInlineShapePicture
        Case "wdInlineShapeLinkedPicture": WdInlineShapeTypeFromString = wdInlineShapeLinkedPicture
        Case "wdInlineShapeOLEControlObject": WdInlineShapeTypeFromString = wdInlineShapeOLEControlObject
        Case "wdInlineShapeHorizontalLine": WdInlineShapeTypeFromString = wdInlineShapeHorizontalLine
        Case "wdInlineShapePictureHorizontalLine": WdInlineShapeTypeFromString = wdInlineShapePictureHorizontalLine
        Case "wdInlineShapeLinkedPictureHorizontalLine": WdInlineShapeTypeFromString = wdInlineShapeLinkedPictureHorizontalLine
        Case "wdInlineShapePictureBullet": WdInlineShapeTypeFromString = wdInlineShapePictureBullet
        Case "wdInlineShapeScriptAnchor": WdInlineShapeTypeFromString = wdInlineShapeScriptAnchor
        Case "wdInlineShapeOWSAnchor": WdInlineShapeTypeFromString = wdInlineShapeOWSAnchor
        Case "wdInlineShapeChart": WdInlineShapeTypeFromString = wdInlineShapeChart
        Case "wdInlineShapeDiagram": WdInlineShapeTypeFromString = wdInlineShapeDiagram
        Case "wdInlineShapeLockedCanvas": WdInlineShapeTypeFromString = wdInlineShapeLockedCanvas
        Case "wdInlineShapeSmartArt": WdInlineShapeTypeFromString = wdInlineShapeSmartArt
    End Select
End Function

Function WdInlineShapeTypeToString(value As WdInlineShapeType) As String
    Select Case value
        Case wdInlineShapeEmbeddedOLEObject: WdInlineShapeTypeToString = "wdInlineShapeEmbeddedOLEObject"
        Case wdInlineShapeLinkedOLEObject: WdInlineShapeTypeToString = "wdInlineShapeLinkedOLEObject"
        Case wdInlineShapePicture: WdInlineShapeTypeToString = "wdInlineShapePicture"
        Case wdInlineShapeLinkedPicture: WdInlineShapeTypeToString = "wdInlineShapeLinkedPicture"
        Case wdInlineShapeOLEControlObject: WdInlineShapeTypeToString = "wdInlineShapeOLEControlObject"
        Case wdInlineShapeHorizontalLine: WdInlineShapeTypeToString = "wdInlineShapeHorizontalLine"
        Case wdInlineShapePictureHorizontalLine: WdInlineShapeTypeToString = "wdInlineShapePictureHorizontalLine"
        Case wdInlineShapeLinkedPictureHorizontalLine: WdInlineShapeTypeToString = "wdInlineShapeLinkedPictureHorizontalLine"
        Case wdInlineShapePictureBullet: WdInlineShapeTypeToString = "wdInlineShapePictureBullet"
        Case wdInlineShapeScriptAnchor: WdInlineShapeTypeToString = "wdInlineShapeScriptAnchor"
        Case wdInlineShapeOWSAnchor: WdInlineShapeTypeToString = "wdInlineShapeOWSAnchor"
        Case wdInlineShapeChart: WdInlineShapeTypeToString = "wdInlineShapeChart"
        Case wdInlineShapeDiagram: WdInlineShapeTypeToString = "wdInlineShapeDiagram"
        Case wdInlineShapeLockedCanvas: WdInlineShapeTypeToString = "wdInlineShapeLockedCanvas"
        Case wdInlineShapeSmartArt: WdInlineShapeTypeToString = "wdInlineShapeSmartArt"
    End Select
End Function
