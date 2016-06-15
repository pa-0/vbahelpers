Attribute VB_Name = "wPbOriginalFormat"
Function PbOriginalFormatFromString(value As String) As PbOriginalFormat
    If IsNumeric(value) Then
        PbOriginalFormatFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbOriginalPublicationFormat": PbOriginalFormatFromString = pbOriginalPublicationFormat
        Case "pbPublisherFile": PbOriginalFormatFromString = pbPublisherFile
    End Select
End Function

Function PbOriginalFormatToString(value As PbOriginalFormat) As String
    Select Case value
        Case pbOriginalPublicationFormat: PbOriginalFormatToString = "pbOriginalPublicationFormat"
        Case pbPublisherFile: PbOriginalFormatToString = "pbPublisherFile"
    End Select
End Function
