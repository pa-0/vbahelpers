Attribute VB_Name = "wSignatureDetail"
Function SignatureDetailFromString(value As String) As SignatureDetail
    If IsNumeric(value) Then
        SignatureDetailFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "sigdetLocalSigningTime": SignatureDetailFromString = sigdetLocalSigningTime
        Case "sigdetApplicationName": SignatureDetailFromString = sigdetApplicationName
        Case "sigdetApplicationVersion": SignatureDetailFromString = sigdetApplicationVersion
        Case "sigdetOfficeVersion": SignatureDetailFromString = sigdetOfficeVersion
        Case "sigdetWindowsVersion": SignatureDetailFromString = sigdetWindowsVersion
        Case "sigdetNumberOfMonitors": SignatureDetailFromString = sigdetNumberOfMonitors
        Case "sigdetHorizResolution": SignatureDetailFromString = sigdetHorizResolution
        Case "sigdetVertResolution": SignatureDetailFromString = sigdetVertResolution
        Case "sigdetColorDepth": SignatureDetailFromString = sigdetColorDepth
        Case "sigdetSignedData": SignatureDetailFromString = sigdetSignedData
        Case "sigdetDocPreviewImg": SignatureDetailFromString = sigdetDocPreviewImg
        Case "sigdetIPFormHash": SignatureDetailFromString = sigdetIPFormHash
        Case "sigdetIPCurrentView": SignatureDetailFromString = sigdetIPCurrentView
        Case "sigdetSignatureType": SignatureDetailFromString = sigdetSignatureType
        Case "sigdetHashAlgorithm": SignatureDetailFromString = sigdetHashAlgorithm
        Case "sigdetShouldShowViewWarning": SignatureDetailFromString = sigdetShouldShowViewWarning
        Case "sigdetDelSuggSigner": SignatureDetailFromString = sigdetDelSuggSigner
        Case "sigdetDelSuggSignerSet": SignatureDetailFromString = sigdetDelSuggSignerSet
        Case "sigdetDelSuggSignerLine2": SignatureDetailFromString = sigdetDelSuggSignerLine2
        Case "sigdetDelSuggSignerLine2Set": SignatureDetailFromString = sigdetDelSuggSignerLine2Set
        Case "sigdetDelSuggSignerEmail": SignatureDetailFromString = sigdetDelSuggSignerEmail
        Case "sigdetDelSuggSignerEmailSet": SignatureDetailFromString = sigdetDelSuggSignerEmailSet
    End Select
End Function

Function SignatureDetailToString(value As SignatureDetail) As String
    Select Case value
        Case sigdetLocalSigningTime: SignatureDetailToString = "sigdetLocalSigningTime"
        Case sigdetApplicationName: SignatureDetailToString = "sigdetApplicationName"
        Case sigdetApplicationVersion: SignatureDetailToString = "sigdetApplicationVersion"
        Case sigdetOfficeVersion: SignatureDetailToString = "sigdetOfficeVersion"
        Case sigdetWindowsVersion: SignatureDetailToString = "sigdetWindowsVersion"
        Case sigdetNumberOfMonitors: SignatureDetailToString = "sigdetNumberOfMonitors"
        Case sigdetHorizResolution: SignatureDetailToString = "sigdetHorizResolution"
        Case sigdetVertResolution: SignatureDetailToString = "sigdetVertResolution"
        Case sigdetColorDepth: SignatureDetailToString = "sigdetColorDepth"
        Case sigdetSignedData: SignatureDetailToString = "sigdetSignedData"
        Case sigdetDocPreviewImg: SignatureDetailToString = "sigdetDocPreviewImg"
        Case sigdetIPFormHash: SignatureDetailToString = "sigdetIPFormHash"
        Case sigdetIPCurrentView: SignatureDetailToString = "sigdetIPCurrentView"
        Case sigdetSignatureType: SignatureDetailToString = "sigdetSignatureType"
        Case sigdetHashAlgorithm: SignatureDetailToString = "sigdetHashAlgorithm"
        Case sigdetShouldShowViewWarning: SignatureDetailToString = "sigdetShouldShowViewWarning"
        Case sigdetDelSuggSigner: SignatureDetailToString = "sigdetDelSuggSigner"
        Case sigdetDelSuggSignerSet: SignatureDetailToString = "sigdetDelSuggSignerSet"
        Case sigdetDelSuggSignerLine2: SignatureDetailToString = "sigdetDelSuggSignerLine2"
        Case sigdetDelSuggSignerLine2Set: SignatureDetailToString = "sigdetDelSuggSignerLine2Set"
        Case sigdetDelSuggSignerEmail: SignatureDetailToString = "sigdetDelSuggSignerEmail"
        Case sigdetDelSuggSignerEmailSet: SignatureDetailToString = "sigdetDelSuggSignerEmailSet"
    End Select
End Function
