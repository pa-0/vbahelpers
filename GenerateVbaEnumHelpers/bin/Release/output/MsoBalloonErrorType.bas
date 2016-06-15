Attribute VB_Name = "wMsoBalloonErrorType"
Function MsoBalloonErrorTypeFromString(value As String) As MsoBalloonErrorType
    If IsNumeric(value) Then
        MsoBalloonErrorTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoBalloonErrorNone": MsoBalloonErrorTypeFromString = msoBalloonErrorNone
        Case "msoBalloonErrorOther": MsoBalloonErrorTypeFromString = msoBalloonErrorOther
        Case "msoBalloonErrorTooBig": MsoBalloonErrorTypeFromString = msoBalloonErrorTooBig
        Case "msoBalloonErrorOutOfMemory": MsoBalloonErrorTypeFromString = msoBalloonErrorOutOfMemory
        Case "msoBalloonErrorBadPictureRef": MsoBalloonErrorTypeFromString = msoBalloonErrorBadPictureRef
        Case "msoBalloonErrorBadReference": MsoBalloonErrorTypeFromString = msoBalloonErrorBadReference
        Case "msoBalloonErrorButtonlessModal": MsoBalloonErrorTypeFromString = msoBalloonErrorButtonlessModal
        Case "msoBalloonErrorButtonModeless": MsoBalloonErrorTypeFromString = msoBalloonErrorButtonModeless
        Case "msoBalloonErrorBadCharacter": MsoBalloonErrorTypeFromString = msoBalloonErrorBadCharacter
        Case "msoBalloonErrorCOMFailure": MsoBalloonErrorTypeFromString = msoBalloonErrorCOMFailure
        Case "msoBalloonErrorCharNotTopmostForModal": MsoBalloonErrorTypeFromString = msoBalloonErrorCharNotTopmostForModal
        Case "msoBalloonErrorTooManyControls": MsoBalloonErrorTypeFromString = msoBalloonErrorTooManyControls
    End Select
End Function

Function MsoBalloonErrorTypeToString(value As MsoBalloonErrorType) As String
    Select Case value
        Case msoBalloonErrorNone: MsoBalloonErrorTypeToString = "msoBalloonErrorNone"
        Case msoBalloonErrorOther: MsoBalloonErrorTypeToString = "msoBalloonErrorOther"
        Case msoBalloonErrorTooBig: MsoBalloonErrorTypeToString = "msoBalloonErrorTooBig"
        Case msoBalloonErrorOutOfMemory: MsoBalloonErrorTypeToString = "msoBalloonErrorOutOfMemory"
        Case msoBalloonErrorBadPictureRef: MsoBalloonErrorTypeToString = "msoBalloonErrorBadPictureRef"
        Case msoBalloonErrorBadReference: MsoBalloonErrorTypeToString = "msoBalloonErrorBadReference"
        Case msoBalloonErrorButtonlessModal: MsoBalloonErrorTypeToString = "msoBalloonErrorButtonlessModal"
        Case msoBalloonErrorButtonModeless: MsoBalloonErrorTypeToString = "msoBalloonErrorButtonModeless"
        Case msoBalloonErrorBadCharacter: MsoBalloonErrorTypeToString = "msoBalloonErrorBadCharacter"
        Case msoBalloonErrorCOMFailure: MsoBalloonErrorTypeToString = "msoBalloonErrorCOMFailure"
        Case msoBalloonErrorCharNotTopmostForModal: MsoBalloonErrorTypeToString = "msoBalloonErrorCharNotTopmostForModal"
        Case msoBalloonErrorTooManyControls: MsoBalloonErrorTypeToString = "msoBalloonErrorTooManyControls"
    End Select
End Function
