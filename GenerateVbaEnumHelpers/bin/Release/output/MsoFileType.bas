Attribute VB_Name = "wMsoFileType"
Function MsoFileTypeFromString(value As String) As MsoFileType
    If IsNumeric(value) Then
        MsoFileTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoFileTypeAllFiles": MsoFileTypeFromString = msoFileTypeAllFiles
        Case "msoFileTypeOfficeFiles": MsoFileTypeFromString = msoFileTypeOfficeFiles
        Case "msoFileTypeWordDocuments": MsoFileTypeFromString = msoFileTypeWordDocuments
        Case "msoFileTypeExcelWorkbooks": MsoFileTypeFromString = msoFileTypeExcelWorkbooks
        Case "msoFileTypePowerPointPresentations": MsoFileTypeFromString = msoFileTypePowerPointPresentations
        Case "msoFileTypeBinders": MsoFileTypeFromString = msoFileTypeBinders
        Case "msoFileTypeDatabases": MsoFileTypeFromString = msoFileTypeDatabases
        Case "msoFileTypeTemplates": MsoFileTypeFromString = msoFileTypeTemplates
        Case "msoFileTypeOutlookItems": MsoFileTypeFromString = msoFileTypeOutlookItems
        Case "msoFileTypeMailItem": MsoFileTypeFromString = msoFileTypeMailItem
        Case "msoFileTypeCalendarItem": MsoFileTypeFromString = msoFileTypeCalendarItem
        Case "msoFileTypeContactItem": MsoFileTypeFromString = msoFileTypeContactItem
        Case "msoFileTypeNoteItem": MsoFileTypeFromString = msoFileTypeNoteItem
        Case "msoFileTypeJournalItem": MsoFileTypeFromString = msoFileTypeJournalItem
        Case "msoFileTypeTaskItem": MsoFileTypeFromString = msoFileTypeTaskItem
        Case "msoFileTypePhotoDrawFiles": MsoFileTypeFromString = msoFileTypePhotoDrawFiles
        Case "msoFileTypeDataConnectionFiles": MsoFileTypeFromString = msoFileTypeDataConnectionFiles
        Case "msoFileTypePublisherFiles": MsoFileTypeFromString = msoFileTypePublisherFiles
        Case "msoFileTypeProjectFiles": MsoFileTypeFromString = msoFileTypeProjectFiles
        Case "msoFileTypeDocumentImagingFiles": MsoFileTypeFromString = msoFileTypeDocumentImagingFiles
        Case "msoFileTypeVisioFiles": MsoFileTypeFromString = msoFileTypeVisioFiles
        Case "msoFileTypeDesignerFiles": MsoFileTypeFromString = msoFileTypeDesignerFiles
        Case "msoFileTypeWebPages": MsoFileTypeFromString = msoFileTypeWebPages
    End Select
End Function

Function MsoFileTypeToString(value As MsoFileType) As String
    Select Case value
        Case msoFileTypeAllFiles: MsoFileTypeToString = "msoFileTypeAllFiles"
        Case msoFileTypeOfficeFiles: MsoFileTypeToString = "msoFileTypeOfficeFiles"
        Case msoFileTypeWordDocuments: MsoFileTypeToString = "msoFileTypeWordDocuments"
        Case msoFileTypeExcelWorkbooks: MsoFileTypeToString = "msoFileTypeExcelWorkbooks"
        Case msoFileTypePowerPointPresentations: MsoFileTypeToString = "msoFileTypePowerPointPresentations"
        Case msoFileTypeBinders: MsoFileTypeToString = "msoFileTypeBinders"
        Case msoFileTypeDatabases: MsoFileTypeToString = "msoFileTypeDatabases"
        Case msoFileTypeTemplates: MsoFileTypeToString = "msoFileTypeTemplates"
        Case msoFileTypeOutlookItems: MsoFileTypeToString = "msoFileTypeOutlookItems"
        Case msoFileTypeMailItem: MsoFileTypeToString = "msoFileTypeMailItem"
        Case msoFileTypeCalendarItem: MsoFileTypeToString = "msoFileTypeCalendarItem"
        Case msoFileTypeContactItem: MsoFileTypeToString = "msoFileTypeContactItem"
        Case msoFileTypeNoteItem: MsoFileTypeToString = "msoFileTypeNoteItem"
        Case msoFileTypeJournalItem: MsoFileTypeToString = "msoFileTypeJournalItem"
        Case msoFileTypeTaskItem: MsoFileTypeToString = "msoFileTypeTaskItem"
        Case msoFileTypePhotoDrawFiles: MsoFileTypeToString = "msoFileTypePhotoDrawFiles"
        Case msoFileTypeDataConnectionFiles: MsoFileTypeToString = "msoFileTypeDataConnectionFiles"
        Case msoFileTypePublisherFiles: MsoFileTypeToString = "msoFileTypePublisherFiles"
        Case msoFileTypeProjectFiles: MsoFileTypeToString = "msoFileTypeProjectFiles"
        Case msoFileTypeDocumentImagingFiles: MsoFileTypeToString = "msoFileTypeDocumentImagingFiles"
        Case msoFileTypeVisioFiles: MsoFileTypeToString = "msoFileTypeVisioFiles"
        Case msoFileTypeDesignerFiles: MsoFileTypeToString = "msoFileTypeDesignerFiles"
        Case msoFileTypeWebPages: MsoFileTypeToString = "msoFileTypeWebPages"
    End Select
End Function
