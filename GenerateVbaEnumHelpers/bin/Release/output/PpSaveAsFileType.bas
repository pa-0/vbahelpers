Attribute VB_Name = "wPpSaveAsFileType"
Function PpSaveAsFileTypeFromString(value As String) As PpSaveAsFileType
    If IsNumeric(value) Then
        PpSaveAsFileTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "ppSaveAsPresentation": PpSaveAsFileTypeFromString = ppSaveAsPresentation
        Case "ppSaveAsPowerPoint7": PpSaveAsFileTypeFromString = ppSaveAsPowerPoint7
        Case "ppSaveAsPowerPoint4": PpSaveAsFileTypeFromString = ppSaveAsPowerPoint4
        Case "ppSaveAsPowerPoint3": PpSaveAsFileTypeFromString = ppSaveAsPowerPoint3
        Case "ppSaveAsTemplate": PpSaveAsFileTypeFromString = ppSaveAsTemplate
        Case "ppSaveAsRTF": PpSaveAsFileTypeFromString = ppSaveAsRTF
        Case "ppSaveAsShow": PpSaveAsFileTypeFromString = ppSaveAsShow
        Case "ppSaveAsAddIn": PpSaveAsFileTypeFromString = ppSaveAsAddIn
        Case "ppSaveAsPowerPoint4FarEast": PpSaveAsFileTypeFromString = ppSaveAsPowerPoint4FarEast
        Case "ppSaveAsDefault": PpSaveAsFileTypeFromString = ppSaveAsDefault
        Case "ppSaveAsHTML": PpSaveAsFileTypeFromString = ppSaveAsHTML
        Case "ppSaveAsHTMLv3": PpSaveAsFileTypeFromString = ppSaveAsHTMLv3
        Case "ppSaveAsHTMLDual": PpSaveAsFileTypeFromString = ppSaveAsHTMLDual
        Case "ppSaveAsMetaFile": PpSaveAsFileTypeFromString = ppSaveAsMetaFile
        Case "ppSaveAsGIF": PpSaveAsFileTypeFromString = ppSaveAsGIF
        Case "ppSaveAsJPG": PpSaveAsFileTypeFromString = ppSaveAsJPG
        Case "ppSaveAsPNG": PpSaveAsFileTypeFromString = ppSaveAsPNG
        Case "ppSaveAsBMP": PpSaveAsFileTypeFromString = ppSaveAsBMP
        Case "ppSaveAsWebArchive": PpSaveAsFileTypeFromString = ppSaveAsWebArchive
        Case "ppSaveAsTIF": PpSaveAsFileTypeFromString = ppSaveAsTIF
        Case "ppSaveAsPresForReview": PpSaveAsFileTypeFromString = ppSaveAsPresForReview
        Case "ppSaveAsEMF": PpSaveAsFileTypeFromString = ppSaveAsEMF
        Case "ppSaveAsOpenXMLPresentation": PpSaveAsFileTypeFromString = ppSaveAsOpenXMLPresentation
        Case "ppSaveAsOpenXMLPresentationMacroEnabled": PpSaveAsFileTypeFromString = ppSaveAsOpenXMLPresentationMacroEnabled
        Case "ppSaveAsOpenXMLTemplate": PpSaveAsFileTypeFromString = ppSaveAsOpenXMLTemplate
        Case "ppSaveAsOpenXMLTemplateMacroEnabled": PpSaveAsFileTypeFromString = ppSaveAsOpenXMLTemplateMacroEnabled
        Case "ppSaveAsOpenXMLShow": PpSaveAsFileTypeFromString = ppSaveAsOpenXMLShow
        Case "ppSaveAsOpenXMLShowMacroEnabled": PpSaveAsFileTypeFromString = ppSaveAsOpenXMLShowMacroEnabled
        Case "ppSaveAsOpenXMLAddin": PpSaveAsFileTypeFromString = ppSaveAsOpenXMLAddin
        Case "ppSaveAsOpenXMLTheme": PpSaveAsFileTypeFromString = ppSaveAsOpenXMLTheme
        Case "ppSaveAsPDF": PpSaveAsFileTypeFromString = ppSaveAsPDF
        Case "ppSaveAsXPS": PpSaveAsFileTypeFromString = ppSaveAsXPS
        Case "ppSaveAsXMLPresentation": PpSaveAsFileTypeFromString = ppSaveAsXMLPresentation
        Case "ppSaveAsOpenDocumentPresentation": PpSaveAsFileTypeFromString = ppSaveAsOpenDocumentPresentation
        Case "ppSaveAsOpenXMLPicturePresentation": PpSaveAsFileTypeFromString = ppSaveAsOpenXMLPicturePresentation
        Case "ppSaveAsWMV": PpSaveAsFileTypeFromString = ppSaveAsWMV
        Case "ppSaveAsExternalConverter": PpSaveAsFileTypeFromString = ppSaveAsExternalConverter
    End Select
End Function

Function PpSaveAsFileTypeToString(value As PpSaveAsFileType) As String
    Select Case value
        Case ppSaveAsPresentation: PpSaveAsFileTypeToString = "ppSaveAsPresentation"
        Case ppSaveAsPowerPoint7: PpSaveAsFileTypeToString = "ppSaveAsPowerPoint7"
        Case ppSaveAsPowerPoint4: PpSaveAsFileTypeToString = "ppSaveAsPowerPoint4"
        Case ppSaveAsPowerPoint3: PpSaveAsFileTypeToString = "ppSaveAsPowerPoint3"
        Case ppSaveAsTemplate: PpSaveAsFileTypeToString = "ppSaveAsTemplate"
        Case ppSaveAsRTF: PpSaveAsFileTypeToString = "ppSaveAsRTF"
        Case ppSaveAsShow: PpSaveAsFileTypeToString = "ppSaveAsShow"
        Case ppSaveAsAddIn: PpSaveAsFileTypeToString = "ppSaveAsAddIn"
        Case ppSaveAsPowerPoint4FarEast: PpSaveAsFileTypeToString = "ppSaveAsPowerPoint4FarEast"
        Case ppSaveAsDefault: PpSaveAsFileTypeToString = "ppSaveAsDefault"
        Case ppSaveAsHTML: PpSaveAsFileTypeToString = "ppSaveAsHTML"
        Case ppSaveAsHTMLv3: PpSaveAsFileTypeToString = "ppSaveAsHTMLv3"
        Case ppSaveAsHTMLDual: PpSaveAsFileTypeToString = "ppSaveAsHTMLDual"
        Case ppSaveAsMetaFile: PpSaveAsFileTypeToString = "ppSaveAsMetaFile"
        Case ppSaveAsGIF: PpSaveAsFileTypeToString = "ppSaveAsGIF"
        Case ppSaveAsJPG: PpSaveAsFileTypeToString = "ppSaveAsJPG"
        Case ppSaveAsPNG: PpSaveAsFileTypeToString = "ppSaveAsPNG"
        Case ppSaveAsBMP: PpSaveAsFileTypeToString = "ppSaveAsBMP"
        Case ppSaveAsWebArchive: PpSaveAsFileTypeToString = "ppSaveAsWebArchive"
        Case ppSaveAsTIF: PpSaveAsFileTypeToString = "ppSaveAsTIF"
        Case ppSaveAsPresForReview: PpSaveAsFileTypeToString = "ppSaveAsPresForReview"
        Case ppSaveAsEMF: PpSaveAsFileTypeToString = "ppSaveAsEMF"
        Case ppSaveAsOpenXMLPresentation: PpSaveAsFileTypeToString = "ppSaveAsOpenXMLPresentation"
        Case ppSaveAsOpenXMLPresentationMacroEnabled: PpSaveAsFileTypeToString = "ppSaveAsOpenXMLPresentationMacroEnabled"
        Case ppSaveAsOpenXMLTemplate: PpSaveAsFileTypeToString = "ppSaveAsOpenXMLTemplate"
        Case ppSaveAsOpenXMLTemplateMacroEnabled: PpSaveAsFileTypeToString = "ppSaveAsOpenXMLTemplateMacroEnabled"
        Case ppSaveAsOpenXMLShow: PpSaveAsFileTypeToString = "ppSaveAsOpenXMLShow"
        Case ppSaveAsOpenXMLShowMacroEnabled: PpSaveAsFileTypeToString = "ppSaveAsOpenXMLShowMacroEnabled"
        Case ppSaveAsOpenXMLAddin: PpSaveAsFileTypeToString = "ppSaveAsOpenXMLAddin"
        Case ppSaveAsOpenXMLTheme: PpSaveAsFileTypeToString = "ppSaveAsOpenXMLTheme"
        Case ppSaveAsPDF: PpSaveAsFileTypeToString = "ppSaveAsPDF"
        Case ppSaveAsXPS: PpSaveAsFileTypeToString = "ppSaveAsXPS"
        Case ppSaveAsXMLPresentation: PpSaveAsFileTypeToString = "ppSaveAsXMLPresentation"
        Case ppSaveAsOpenDocumentPresentation: PpSaveAsFileTypeToString = "ppSaveAsOpenDocumentPresentation"
        Case ppSaveAsOpenXMLPicturePresentation: PpSaveAsFileTypeToString = "ppSaveAsOpenXMLPicturePresentation"
        Case ppSaveAsWMV: PpSaveAsFileTypeToString = "ppSaveAsWMV"
        Case ppSaveAsExternalConverter: PpSaveAsFileTypeToString = "ppSaveAsExternalConverter"
    End Select
End Function
