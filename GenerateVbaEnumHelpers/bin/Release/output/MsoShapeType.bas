Attribute VB_Name = "wMsoShapeType"
Function MsoShapeTypeFromString(value As String) As MsoShapeType
    If IsNumeric(value) Then
        MsoShapeTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoAutoShape": MsoShapeTypeFromString = msoAutoShape
        Case "msoCallout": MsoShapeTypeFromString = msoCallout
        Case "msoChart": MsoShapeTypeFromString = msoChart
        Case "msoComment": MsoShapeTypeFromString = msoComment
        Case "msoFreeform": MsoShapeTypeFromString = msoFreeform
        Case "msoGroup": MsoShapeTypeFromString = msoGroup
        Case "msoEmbeddedOLEObject": MsoShapeTypeFromString = msoEmbeddedOLEObject
        Case "msoFormControl": MsoShapeTypeFromString = msoFormControl
        Case "msoLine": MsoShapeTypeFromString = msoLine
        Case "msoLinkedOLEObject": MsoShapeTypeFromString = msoLinkedOLEObject
        Case "msoLinkedPicture": MsoShapeTypeFromString = msoLinkedPicture
        Case "msoOLEControlObject": MsoShapeTypeFromString = msoOLEControlObject
        Case "msoPicture": MsoShapeTypeFromString = msoPicture
        Case "msoPlaceholder": MsoShapeTypeFromString = msoPlaceholder
        Case "msoTextEffect": MsoShapeTypeFromString = msoTextEffect
        Case "msoMedia": MsoShapeTypeFromString = msoMedia
        Case "msoTextBox": MsoShapeTypeFromString = msoTextBox
        Case "msoScriptAnchor": MsoShapeTypeFromString = msoScriptAnchor
        Case "msoTable": MsoShapeTypeFromString = msoTable
        Case "msoCanvas": MsoShapeTypeFromString = msoCanvas
        Case "msoDiagram": MsoShapeTypeFromString = msoDiagram
        Case "msoInk": MsoShapeTypeFromString = msoInk
        Case "msoInkComment": MsoShapeTypeFromString = msoInkComment
        Case "msoSmartArt": MsoShapeTypeFromString = msoSmartArt
        Case "msoSlicer": MsoShapeTypeFromString = msoSlicer
        Case "msoShapeTypeMixed": MsoShapeTypeFromString = msoShapeTypeMixed
    End Select
End Function

Function MsoShapeTypeToString(value As MsoShapeType) As String
    Select Case value
        Case msoAutoShape: MsoShapeTypeToString = "msoAutoShape"
        Case msoCallout: MsoShapeTypeToString = "msoCallout"
        Case msoChart: MsoShapeTypeToString = "msoChart"
        Case msoComment: MsoShapeTypeToString = "msoComment"
        Case msoFreeform: MsoShapeTypeToString = "msoFreeform"
        Case msoGroup: MsoShapeTypeToString = "msoGroup"
        Case msoEmbeddedOLEObject: MsoShapeTypeToString = "msoEmbeddedOLEObject"
        Case msoFormControl: MsoShapeTypeToString = "msoFormControl"
        Case msoLine: MsoShapeTypeToString = "msoLine"
        Case msoLinkedOLEObject: MsoShapeTypeToString = "msoLinkedOLEObject"
        Case msoLinkedPicture: MsoShapeTypeToString = "msoLinkedPicture"
        Case msoOLEControlObject: MsoShapeTypeToString = "msoOLEControlObject"
        Case msoPicture: MsoShapeTypeToString = "msoPicture"
        Case msoPlaceholder: MsoShapeTypeToString = "msoPlaceholder"
        Case msoTextEffect: MsoShapeTypeToString = "msoTextEffect"
        Case msoMedia: MsoShapeTypeToString = "msoMedia"
        Case msoTextBox: MsoShapeTypeToString = "msoTextBox"
        Case msoScriptAnchor: MsoShapeTypeToString = "msoScriptAnchor"
        Case msoTable: MsoShapeTypeToString = "msoTable"
        Case msoCanvas: MsoShapeTypeToString = "msoCanvas"
        Case msoDiagram: MsoShapeTypeToString = "msoDiagram"
        Case msoInk: MsoShapeTypeToString = "msoInk"
        Case msoInkComment: MsoShapeTypeToString = "msoInkComment"
        Case msoSmartArt: MsoShapeTypeToString = "msoSmartArt"
        Case msoSlicer: MsoShapeTypeToString = "msoSlicer"
        Case msoShapeTypeMixed: MsoShapeTypeToString = "msoShapeTypeMixed"
    End Select
End Function
