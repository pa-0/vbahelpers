Attribute VB_Name = "wWdRemoveDocInfoType"
Function WdRemoveDocInfoTypeFromString(value As String) As WdRemoveDocInfoType
    If IsNumeric(value) Then
        WdRemoveDocInfoTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdRDIComments": WdRemoveDocInfoTypeFromString = wdRDIComments
        Case "wdRDIRevisions": WdRemoveDocInfoTypeFromString = wdRDIRevisions
        Case "wdRDIVersions": WdRemoveDocInfoTypeFromString = wdRDIVersions
        Case "wdRDIRemovePersonalInformation": WdRemoveDocInfoTypeFromString = wdRDIRemovePersonalInformation
        Case "wdRDIEmailHeader": WdRemoveDocInfoTypeFromString = wdRDIEmailHeader
        Case "wdRDIRoutingSlip": WdRemoveDocInfoTypeFromString = wdRDIRoutingSlip
        Case "wdRDISendForReview": WdRemoveDocInfoTypeFromString = wdRDISendForReview
        Case "wdRDIDocumentProperties": WdRemoveDocInfoTypeFromString = wdRDIDocumentProperties
        Case "wdRDITemplate": WdRemoveDocInfoTypeFromString = wdRDITemplate
        Case "wdRDIDocumentWorkspace": WdRemoveDocInfoTypeFromString = wdRDIDocumentWorkspace
        Case "wdRDIInkAnnotations": WdRemoveDocInfoTypeFromString = wdRDIInkAnnotations
        Case "wdRDIDocumentServerProperties": WdRemoveDocInfoTypeFromString = wdRDIDocumentServerProperties
        Case "wdRDIDocumentManagementPolicy": WdRemoveDocInfoTypeFromString = wdRDIDocumentManagementPolicy
        Case "wdRDIContentType": WdRemoveDocInfoTypeFromString = wdRDIContentType
        Case "wdRDIAll": WdRemoveDocInfoTypeFromString = wdRDIAll
    End Select
End Function

Function WdRemoveDocInfoTypeToString(value As WdRemoveDocInfoType) As String
    Select Case value
        Case wdRDIComments: WdRemoveDocInfoTypeToString = "wdRDIComments"
        Case wdRDIRevisions: WdRemoveDocInfoTypeToString = "wdRDIRevisions"
        Case wdRDIVersions: WdRemoveDocInfoTypeToString = "wdRDIVersions"
        Case wdRDIRemovePersonalInformation: WdRemoveDocInfoTypeToString = "wdRDIRemovePersonalInformation"
        Case wdRDIEmailHeader: WdRemoveDocInfoTypeToString = "wdRDIEmailHeader"
        Case wdRDIRoutingSlip: WdRemoveDocInfoTypeToString = "wdRDIRoutingSlip"
        Case wdRDISendForReview: WdRemoveDocInfoTypeToString = "wdRDISendForReview"
        Case wdRDIDocumentProperties: WdRemoveDocInfoTypeToString = "wdRDIDocumentProperties"
        Case wdRDITemplate: WdRemoveDocInfoTypeToString = "wdRDITemplate"
        Case wdRDIDocumentWorkspace: WdRemoveDocInfoTypeToString = "wdRDIDocumentWorkspace"
        Case wdRDIInkAnnotations: WdRemoveDocInfoTypeToString = "wdRDIInkAnnotations"
        Case wdRDIDocumentServerProperties: WdRemoveDocInfoTypeToString = "wdRDIDocumentServerProperties"
        Case wdRDIDocumentManagementPolicy: WdRemoveDocInfoTypeToString = "wdRDIDocumentManagementPolicy"
        Case wdRDIContentType: WdRemoveDocInfoTypeToString = "wdRDIContentType"
        Case wdRDIAll: WdRemoveDocInfoTypeToString = "wdRDIAll"
    End Select
End Function
