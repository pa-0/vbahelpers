Attribute VB_Name = "wPbShapeType"
Function PbShapeTypeFromString(value As String) As PbShapeType
    If IsNumeric(value) Then
        PbShapeTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbAutoShape": PbShapeTypeFromString = pbAutoShape
        Case "pbCallout": PbShapeTypeFromString = pbCallout
        Case "pbChart": PbShapeTypeFromString = pbChart
        Case "pbComment": PbShapeTypeFromString = pbComment
        Case "pbFreeform": PbShapeTypeFromString = pbFreeform
        Case "pbGroup": PbShapeTypeFromString = pbGroup
        Case "pbEmbeddedOLEObject": PbShapeTypeFromString = pbEmbeddedOLEObject
        Case "pbFormControl": PbShapeTypeFromString = pbFormControl
        Case "pbLine": PbShapeTypeFromString = pbLine
        Case "pbLinkedOLEObject": PbShapeTypeFromString = pbLinkedOLEObject
        Case "pbLinkedPicture": PbShapeTypeFromString = pbLinkedPicture
        Case "pbOLEControlObject": PbShapeTypeFromString = pbOLEControlObject
        Case "pbPicture": PbShapeTypeFromString = pbPicture
        Case "pbPlaceholder": PbShapeTypeFromString = pbPlaceholder
        Case "pbTextEffect": PbShapeTypeFromString = pbTextEffect
        Case "pbMedia": PbShapeTypeFromString = pbMedia
        Case "pbTextFrame": PbShapeTypeFromString = pbTextFrame
        Case "pbTable": PbShapeTypeFromString = pbTable
        Case "pbWebCheckBox": PbShapeTypeFromString = pbWebCheckBox
        Case "pbWebCommandButton": PbShapeTypeFromString = pbWebCommandButton
        Case "pbWebListBox": PbShapeTypeFromString = pbWebListBox
        Case "pbWebMultiLineTextBox": PbShapeTypeFromString = pbWebMultiLineTextBox
        Case "pbWebOptionButton": PbShapeTypeFromString = pbWebOptionButton
        Case "pbWebSingleLineTextBox": PbShapeTypeFromString = pbWebSingleLineTextBox
        Case "pbWebWebComponent": PbShapeTypeFromString = pbWebWebComponent
        Case "pbWebHTMLFragment": PbShapeTypeFromString = pbWebHTMLFragment
        Case "pbGroupWizard": PbShapeTypeFromString = pbGroupWizard
        Case "pbWebHotSpot": PbShapeTypeFromString = pbWebHotSpot
        Case "pbCatalogMergeArea": PbShapeTypeFromString = pbCatalogMergeArea
        Case "pbWebNavigationBar": PbShapeTypeFromString = pbWebNavigationBar
        Case "pbBarCodePictureHolder": PbShapeTypeFromString = pbBarCodePictureHolder
        Case "pbShapeTypeMixed": PbShapeTypeFromString = pbShapeTypeMixed
    End Select
End Function

Function PbShapeTypeToString(value As PbShapeType) As String
    Select Case value
        Case pbAutoShape: PbShapeTypeToString = "pbAutoShape"
        Case pbCallout: PbShapeTypeToString = "pbCallout"
        Case pbChart: PbShapeTypeToString = "pbChart"
        Case pbComment: PbShapeTypeToString = "pbComment"
        Case pbFreeform: PbShapeTypeToString = "pbFreeform"
        Case pbGroup: PbShapeTypeToString = "pbGroup"
        Case pbEmbeddedOLEObject: PbShapeTypeToString = "pbEmbeddedOLEObject"
        Case pbFormControl: PbShapeTypeToString = "pbFormControl"
        Case pbLine: PbShapeTypeToString = "pbLine"
        Case pbLinkedOLEObject: PbShapeTypeToString = "pbLinkedOLEObject"
        Case pbLinkedPicture: PbShapeTypeToString = "pbLinkedPicture"
        Case pbOLEControlObject: PbShapeTypeToString = "pbOLEControlObject"
        Case pbPicture: PbShapeTypeToString = "pbPicture"
        Case pbPlaceholder: PbShapeTypeToString = "pbPlaceholder"
        Case pbTextEffect: PbShapeTypeToString = "pbTextEffect"
        Case pbMedia: PbShapeTypeToString = "pbMedia"
        Case pbTextFrame: PbShapeTypeToString = "pbTextFrame"
        Case pbTable: PbShapeTypeToString = "pbTable"
        Case pbWebCheckBox: PbShapeTypeToString = "pbWebCheckBox"
        Case pbWebCommandButton: PbShapeTypeToString = "pbWebCommandButton"
        Case pbWebListBox: PbShapeTypeToString = "pbWebListBox"
        Case pbWebMultiLineTextBox: PbShapeTypeToString = "pbWebMultiLineTextBox"
        Case pbWebOptionButton: PbShapeTypeToString = "pbWebOptionButton"
        Case pbWebSingleLineTextBox: PbShapeTypeToString = "pbWebSingleLineTextBox"
        Case pbWebWebComponent: PbShapeTypeToString = "pbWebWebComponent"
        Case pbWebHTMLFragment: PbShapeTypeToString = "pbWebHTMLFragment"
        Case pbGroupWizard: PbShapeTypeToString = "pbGroupWizard"
        Case pbWebHotSpot: PbShapeTypeToString = "pbWebHotSpot"
        Case pbCatalogMergeArea: PbShapeTypeToString = "pbCatalogMergeArea"
        Case pbWebNavigationBar: PbShapeTypeToString = "pbWebNavigationBar"
        Case pbBarCodePictureHolder: PbShapeTypeToString = "pbBarCodePictureHolder"
        Case pbShapeTypeMixed: PbShapeTypeToString = "pbShapeTypeMixed"
    End Select
End Function
