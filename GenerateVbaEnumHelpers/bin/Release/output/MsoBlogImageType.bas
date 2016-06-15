Attribute VB_Name = "wMsoBlogImageType"
Function MsoBlogImageTypeFromString(value As String) As MsoBlogImageType
    If IsNumeric(value) Then
        MsoBlogImageTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoblogImageTypeJPEG": MsoBlogImageTypeFromString = msoblogImageTypeJPEG
        Case "msoblogImageTypeGIF": MsoBlogImageTypeFromString = msoblogImageTypeGIF
        Case "msoblogImageTypePNG": MsoBlogImageTypeFromString = msoblogImageTypePNG
    End Select
End Function

Function MsoBlogImageTypeToString(value As MsoBlogImageType) As String
    Select Case value
        Case msoblogImageTypeJPEG: MsoBlogImageTypeToString = "msoblogImageTypeJPEG"
        Case msoblogImageTypeGIF: MsoBlogImageTypeToString = "msoblogImageTypeGIF"
        Case msoblogImageTypePNG: MsoBlogImageTypeToString = "msoblogImageTypePNG"
    End Select
End Function
