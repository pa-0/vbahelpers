Attribute VB_Name = "wMsoSyncErrorType"
Function MsoSyncErrorTypeFromString(value As String) As MsoSyncErrorType
    If IsNumeric(value) Then
        MsoSyncErrorTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoSyncErrorNone": MsoSyncErrorTypeFromString = msoSyncErrorNone
        Case "msoSyncErrorUnauthorizedUser": MsoSyncErrorTypeFromString = msoSyncErrorUnauthorizedUser
        Case "msoSyncErrorCouldNotConnect": MsoSyncErrorTypeFromString = msoSyncErrorCouldNotConnect
        Case "msoSyncErrorOutOfSpace": MsoSyncErrorTypeFromString = msoSyncErrorOutOfSpace
        Case "msoSyncErrorFileNotFound": MsoSyncErrorTypeFromString = msoSyncErrorFileNotFound
        Case "msoSyncErrorFileTooLarge": MsoSyncErrorTypeFromString = msoSyncErrorFileTooLarge
        Case "msoSyncErrorFileInUse": MsoSyncErrorTypeFromString = msoSyncErrorFileInUse
        Case "msoSyncErrorVirusUpload": MsoSyncErrorTypeFromString = msoSyncErrorVirusUpload
        Case "msoSyncErrorVirusDownload": MsoSyncErrorTypeFromString = msoSyncErrorVirusDownload
        Case "msoSyncErrorUnknownUpload": MsoSyncErrorTypeFromString = msoSyncErrorUnknownUpload
        Case "msoSyncErrorUnknownDownload": MsoSyncErrorTypeFromString = msoSyncErrorUnknownDownload
        Case "msoSyncErrorCouldNotOpen": MsoSyncErrorTypeFromString = msoSyncErrorCouldNotOpen
        Case "msoSyncErrorCouldNotUpdate": MsoSyncErrorTypeFromString = msoSyncErrorCouldNotUpdate
        Case "msoSyncErrorCouldNotCompare": MsoSyncErrorTypeFromString = msoSyncErrorCouldNotCompare
        Case "msoSyncErrorCouldNotResolve": MsoSyncErrorTypeFromString = msoSyncErrorCouldNotResolve
        Case "msoSyncErrorNoNetwork": MsoSyncErrorTypeFromString = msoSyncErrorNoNetwork
        Case "msoSyncErrorUnknown": MsoSyncErrorTypeFromString = msoSyncErrorUnknown
    End Select
End Function

Function MsoSyncErrorTypeToString(value As MsoSyncErrorType) As String
    Select Case value
        Case msoSyncErrorNone: MsoSyncErrorTypeToString = "msoSyncErrorNone"
        Case msoSyncErrorUnauthorizedUser: MsoSyncErrorTypeToString = "msoSyncErrorUnauthorizedUser"
        Case msoSyncErrorCouldNotConnect: MsoSyncErrorTypeToString = "msoSyncErrorCouldNotConnect"
        Case msoSyncErrorOutOfSpace: MsoSyncErrorTypeToString = "msoSyncErrorOutOfSpace"
        Case msoSyncErrorFileNotFound: MsoSyncErrorTypeToString = "msoSyncErrorFileNotFound"
        Case msoSyncErrorFileTooLarge: MsoSyncErrorTypeToString = "msoSyncErrorFileTooLarge"
        Case msoSyncErrorFileInUse: MsoSyncErrorTypeToString = "msoSyncErrorFileInUse"
        Case msoSyncErrorVirusUpload: MsoSyncErrorTypeToString = "msoSyncErrorVirusUpload"
        Case msoSyncErrorVirusDownload: MsoSyncErrorTypeToString = "msoSyncErrorVirusDownload"
        Case msoSyncErrorUnknownUpload: MsoSyncErrorTypeToString = "msoSyncErrorUnknownUpload"
        Case msoSyncErrorUnknownDownload: MsoSyncErrorTypeToString = "msoSyncErrorUnknownDownload"
        Case msoSyncErrorCouldNotOpen: MsoSyncErrorTypeToString = "msoSyncErrorCouldNotOpen"
        Case msoSyncErrorCouldNotUpdate: MsoSyncErrorTypeToString = "msoSyncErrorCouldNotUpdate"
        Case msoSyncErrorCouldNotCompare: MsoSyncErrorTypeToString = "msoSyncErrorCouldNotCompare"
        Case msoSyncErrorCouldNotResolve: MsoSyncErrorTypeToString = "msoSyncErrorCouldNotResolve"
        Case msoSyncErrorNoNetwork: MsoSyncErrorTypeToString = "msoSyncErrorNoNetwork"
        Case msoSyncErrorUnknown: MsoSyncErrorTypeToString = "msoSyncErrorUnknown"
    End Select
End Function
