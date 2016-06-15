Attribute VB_Name = "wPbFileFormat"
Function PbFileFormatFromString(value As String) As PbFileFormat
    If IsNumeric(value) Then
        PbFileFormatFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbFilePublication": PbFileFormatFromString = pbFilePublication
        Case "pbFilePublisher98": PbFileFormatFromString = pbFilePublisher98
        Case "pbFilePublisher2000": PbFileFormatFromString = pbFilePublisher2000
        Case "pbFilePublicationHTML": PbFileFormatFromString = pbFilePublicationHTML
        Case "pbFileWebArchive": PbFileFormatFromString = pbFileWebArchive
        Case "pbFileRTF": PbFileFormatFromString = pbFileRTF
        Case "pbFileHTMLFiltered": PbFileFormatFromString = pbFileHTMLFiltered
        Case "pbFilePlainText": PbFileFormatFromString = pbFilePlainText
        Case "pbFileUnicodeText": PbFileFormatFromString = pbFileUnicodeText
    End Select
End Function

Function PbFileFormatToString(value As PbFileFormat) As String
    Select Case value
        Case pbFilePublication: PbFileFormatToString = "pbFilePublication"
        Case pbFilePublisher98: PbFileFormatToString = "pbFilePublisher98"
        Case pbFilePublisher2000: PbFileFormatToString = "pbFilePublisher2000"
        Case pbFilePublicationHTML: PbFileFormatToString = "pbFilePublicationHTML"
        Case pbFileWebArchive: PbFileFormatToString = "pbFileWebArchive"
        Case pbFileRTF: PbFileFormatToString = "pbFileRTF"
        Case pbFileHTMLFiltered: PbFileFormatToString = "pbFileHTMLFiltered"
        Case pbFilePlainText: PbFileFormatToString = "pbFilePlainText"
        Case pbFileUnicodeText: PbFileFormatToString = "pbFileUnicodeText"
    End Select
End Function
