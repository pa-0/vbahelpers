Attribute VB_Name = "wXlRemoveDocInfoType"
Function XlRemoveDocInfoTypeFromString(value As String) As XlRemoveDocInfoType
    If IsNumeric(value) Then
        XlRemoveDocInfoTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlRDIComments": XlRemoveDocInfoTypeFromString = xlRDIComments
        Case "xlRDIRemovePersonalInformation": XlRemoveDocInfoTypeFromString = xlRDIRemovePersonalInformation
        Case "xlRDIEmailHeader": XlRemoveDocInfoTypeFromString = xlRDIEmailHeader
        Case "xlRDIRoutingSlip": XlRemoveDocInfoTypeFromString = xlRDIRoutingSlip
        Case "xlRDISendForReview": XlRemoveDocInfoTypeFromString = xlRDISendForReview
        Case "xlRDIDocumentProperties": XlRemoveDocInfoTypeFromString = xlRDIDocumentProperties
        Case "xlRDIDocumentWorkspace": XlRemoveDocInfoTypeFromString = xlRDIDocumentWorkspace
        Case "xlRDIInkAnnotations": XlRemoveDocInfoTypeFromString = xlRDIInkAnnotations
        Case "xlRDIScenarioComments": XlRemoveDocInfoTypeFromString = xlRDIScenarioComments
        Case "xlRDIPublishInfo": XlRemoveDocInfoTypeFromString = xlRDIPublishInfo
        Case "xlRDIDocumentServerProperties": XlRemoveDocInfoTypeFromString = xlRDIDocumentServerProperties
        Case "xlRDIDocumentManagementPolicy": XlRemoveDocInfoTypeFromString = xlRDIDocumentManagementPolicy
        Case "xlRDIContentType": XlRemoveDocInfoTypeFromString = xlRDIContentType
        Case "xlRDIDefinedNameComments": XlRemoveDocInfoTypeFromString = xlRDIDefinedNameComments
        Case "xlRDIInactiveDataConnections": XlRemoveDocInfoTypeFromString = xlRDIInactiveDataConnections
        Case "xlRDIPrinterPath": XlRemoveDocInfoTypeFromString = xlRDIPrinterPath
        Case "xlRDIAll": XlRemoveDocInfoTypeFromString = xlRDIAll
    End Select
End Function

Function XlRemoveDocInfoTypeToString(value As XlRemoveDocInfoType) As String
    Select Case value
        Case xlRDIComments: XlRemoveDocInfoTypeToString = "xlRDIComments"
        Case xlRDIRemovePersonalInformation: XlRemoveDocInfoTypeToString = "xlRDIRemovePersonalInformation"
        Case xlRDIEmailHeader: XlRemoveDocInfoTypeToString = "xlRDIEmailHeader"
        Case xlRDIRoutingSlip: XlRemoveDocInfoTypeToString = "xlRDIRoutingSlip"
        Case xlRDISendForReview: XlRemoveDocInfoTypeToString = "xlRDISendForReview"
        Case xlRDIDocumentProperties: XlRemoveDocInfoTypeToString = "xlRDIDocumentProperties"
        Case xlRDIDocumentWorkspace: XlRemoveDocInfoTypeToString = "xlRDIDocumentWorkspace"
        Case xlRDIInkAnnotations: XlRemoveDocInfoTypeToString = "xlRDIInkAnnotations"
        Case xlRDIScenarioComments: XlRemoveDocInfoTypeToString = "xlRDIScenarioComments"
        Case xlRDIPublishInfo: XlRemoveDocInfoTypeToString = "xlRDIPublishInfo"
        Case xlRDIDocumentServerProperties: XlRemoveDocInfoTypeToString = "xlRDIDocumentServerProperties"
        Case xlRDIDocumentManagementPolicy: XlRemoveDocInfoTypeToString = "xlRDIDocumentManagementPolicy"
        Case xlRDIContentType: XlRemoveDocInfoTypeToString = "xlRDIContentType"
        Case xlRDIDefinedNameComments: XlRemoveDocInfoTypeToString = "xlRDIDefinedNameComments"
        Case xlRDIInactiveDataConnections: XlRemoveDocInfoTypeToString = "xlRDIInactiveDataConnections"
        Case xlRDIPrinterPath: XlRemoveDocInfoTypeToString = "xlRDIPrinterPath"
        Case xlRDIAll: XlRemoveDocInfoTypeToString = "xlRDIAll"
    End Select
End Function
