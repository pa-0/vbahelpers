Attribute VB_Name = "wPpRemoveDocInfoType"
Function PpRemoveDocInfoTypeFromString(value As String) As PpRemoveDocInfoType
    If IsNumeric(value) Then
        PpRemoveDocInfoTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "ppRDIComments": PpRemoveDocInfoTypeFromString = ppRDIComments
        Case "ppRDIRemovePersonalInformation": PpRemoveDocInfoTypeFromString = ppRDIRemovePersonalInformation
        Case "ppRDIDocumentProperties": PpRemoveDocInfoTypeFromString = ppRDIDocumentProperties
        Case "ppRDIDocumentWorkspace": PpRemoveDocInfoTypeFromString = ppRDIDocumentWorkspace
        Case "ppRDIInkAnnotations": PpRemoveDocInfoTypeFromString = ppRDIInkAnnotations
        Case "ppRDIPublishPath": PpRemoveDocInfoTypeFromString = ppRDIPublishPath
        Case "ppRDIDocumentServerProperties": PpRemoveDocInfoTypeFromString = ppRDIDocumentServerProperties
        Case "ppRDIDocumentManagementPolicy": PpRemoveDocInfoTypeFromString = ppRDIDocumentManagementPolicy
        Case "ppRDIContentType": PpRemoveDocInfoTypeFromString = ppRDIContentType
        Case "ppRDISlideUpdateInformation": PpRemoveDocInfoTypeFromString = ppRDISlideUpdateInformation
        Case "ppRDIAll": PpRemoveDocInfoTypeFromString = ppRDIAll
    End Select
End Function

Function PpRemoveDocInfoTypeToString(value As PpRemoveDocInfoType) As String
    Select Case value
        Case ppRDIComments: PpRemoveDocInfoTypeToString = "ppRDIComments"
        Case ppRDIRemovePersonalInformation: PpRemoveDocInfoTypeToString = "ppRDIRemovePersonalInformation"
        Case ppRDIDocumentProperties: PpRemoveDocInfoTypeToString = "ppRDIDocumentProperties"
        Case ppRDIDocumentWorkspace: PpRemoveDocInfoTypeToString = "ppRDIDocumentWorkspace"
        Case ppRDIInkAnnotations: PpRemoveDocInfoTypeToString = "ppRDIInkAnnotations"
        Case ppRDIPublishPath: PpRemoveDocInfoTypeToString = "ppRDIPublishPath"
        Case ppRDIDocumentServerProperties: PpRemoveDocInfoTypeToString = "ppRDIDocumentServerProperties"
        Case ppRDIDocumentManagementPolicy: PpRemoveDocInfoTypeToString = "ppRDIDocumentManagementPolicy"
        Case ppRDIContentType: PpRemoveDocInfoTypeToString = "ppRDIContentType"
        Case ppRDISlideUpdateInformation: PpRemoveDocInfoTypeToString = "ppRDISlideUpdateInformation"
        Case ppRDIAll: PpRemoveDocInfoTypeToString = "ppRDIAll"
    End Select
End Function
