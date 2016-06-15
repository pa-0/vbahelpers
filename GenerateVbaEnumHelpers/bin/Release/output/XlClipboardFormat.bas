Attribute VB_Name = "wXlClipboardFormat"
Function XlClipboardFormatFromString(value As String) As XlClipboardFormat
    If IsNumeric(value) Then
        XlClipboardFormatFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlClipboardFormatText": XlClipboardFormatFromString = xlClipboardFormatText
        Case "xlClipboardFormatVALU": XlClipboardFormatFromString = xlClipboardFormatVALU
        Case "xlClipboardFormatPICT": XlClipboardFormatFromString = xlClipboardFormatPICT
        Case "xlClipboardFormatPrintPICT": XlClipboardFormatFromString = xlClipboardFormatPrintPICT
        Case "xlClipboardFormatDIF": XlClipboardFormatFromString = xlClipboardFormatDIF
        Case "xlClipboardFormatCSV": XlClipboardFormatFromString = xlClipboardFormatCSV
        Case "xlClipboardFormatSYLK": XlClipboardFormatFromString = xlClipboardFormatSYLK
        Case "xlClipboardFormatRTF": XlClipboardFormatFromString = xlClipboardFormatRTF
        Case "xlClipboardFormatBIFF": XlClipboardFormatFromString = xlClipboardFormatBIFF
        Case "xlClipboardFormatBitmap": XlClipboardFormatFromString = xlClipboardFormatBitmap
        Case "xlClipboardFormatWK1": XlClipboardFormatFromString = xlClipboardFormatWK1
        Case "xlClipboardFormatLink": XlClipboardFormatFromString = xlClipboardFormatLink
        Case "xlClipboardFormatDspText": XlClipboardFormatFromString = xlClipboardFormatDspText
        Case "xlClipboardFormatCGM": XlClipboardFormatFromString = xlClipboardFormatCGM
        Case "xlClipboardFormatNative": XlClipboardFormatFromString = xlClipboardFormatNative
        Case "xlClipboardFormatBinary": XlClipboardFormatFromString = xlClipboardFormatBinary
        Case "xlClipboardFormatTable": XlClipboardFormatFromString = xlClipboardFormatTable
        Case "xlClipboardFormatOwnerLink": XlClipboardFormatFromString = xlClipboardFormatOwnerLink
        Case "xlClipboardFormatBIFF2": XlClipboardFormatFromString = xlClipboardFormatBIFF2
        Case "xlClipboardFormatObjectLink": XlClipboardFormatFromString = xlClipboardFormatObjectLink
        Case "xlClipboardFormatBIFF3": XlClipboardFormatFromString = xlClipboardFormatBIFF3
        Case "xlClipboardFormatEmbeddedObject": XlClipboardFormatFromString = xlClipboardFormatEmbeddedObject
        Case "xlClipboardFormatEmbedSource": XlClipboardFormatFromString = xlClipboardFormatEmbedSource
        Case "xlClipboardFormatLinkSource": XlClipboardFormatFromString = xlClipboardFormatLinkSource
        Case "xlClipboardFormatMovie": XlClipboardFormatFromString = xlClipboardFormatMovie
        Case "xlClipboardFormatToolFace": XlClipboardFormatFromString = xlClipboardFormatToolFace
        Case "xlClipboardFormatToolFacePICT": XlClipboardFormatFromString = xlClipboardFormatToolFacePICT
        Case "xlClipboardFormatStandardScale": XlClipboardFormatFromString = xlClipboardFormatStandardScale
        Case "xlClipboardFormatStandardFont": XlClipboardFormatFromString = xlClipboardFormatStandardFont
        Case "xlClipboardFormatScreenPICT": XlClipboardFormatFromString = xlClipboardFormatScreenPICT
        Case "xlClipboardFormatBIFF4": XlClipboardFormatFromString = xlClipboardFormatBIFF4
        Case "xlClipboardFormatObjectDesc": XlClipboardFormatFromString = xlClipboardFormatObjectDesc
        Case "xlClipboardFormatLinkSourceDesc": XlClipboardFormatFromString = xlClipboardFormatLinkSourceDesc
        Case "xlClipboardFormatBIFF12": XlClipboardFormatFromString = xlClipboardFormatBIFF12
    End Select
End Function

Function XlClipboardFormatToString(value As XlClipboardFormat) As String
    Select Case value
        Case xlClipboardFormatText: XlClipboardFormatToString = "xlClipboardFormatText"
        Case xlClipboardFormatVALU: XlClipboardFormatToString = "xlClipboardFormatVALU"
        Case xlClipboardFormatPICT: XlClipboardFormatToString = "xlClipboardFormatPICT"
        Case xlClipboardFormatPrintPICT: XlClipboardFormatToString = "xlClipboardFormatPrintPICT"
        Case xlClipboardFormatDIF: XlClipboardFormatToString = "xlClipboardFormatDIF"
        Case xlClipboardFormatCSV: XlClipboardFormatToString = "xlClipboardFormatCSV"
        Case xlClipboardFormatSYLK: XlClipboardFormatToString = "xlClipboardFormatSYLK"
        Case xlClipboardFormatRTF: XlClipboardFormatToString = "xlClipboardFormatRTF"
        Case xlClipboardFormatBIFF: XlClipboardFormatToString = "xlClipboardFormatBIFF"
        Case xlClipboardFormatBitmap: XlClipboardFormatToString = "xlClipboardFormatBitmap"
        Case xlClipboardFormatWK1: XlClipboardFormatToString = "xlClipboardFormatWK1"
        Case xlClipboardFormatLink: XlClipboardFormatToString = "xlClipboardFormatLink"
        Case xlClipboardFormatDspText: XlClipboardFormatToString = "xlClipboardFormatDspText"
        Case xlClipboardFormatCGM: XlClipboardFormatToString = "xlClipboardFormatCGM"
        Case xlClipboardFormatNative: XlClipboardFormatToString = "xlClipboardFormatNative"
        Case xlClipboardFormatBinary: XlClipboardFormatToString = "xlClipboardFormatBinary"
        Case xlClipboardFormatTable: XlClipboardFormatToString = "xlClipboardFormatTable"
        Case xlClipboardFormatOwnerLink: XlClipboardFormatToString = "xlClipboardFormatOwnerLink"
        Case xlClipboardFormatBIFF2: XlClipboardFormatToString = "xlClipboardFormatBIFF2"
        Case xlClipboardFormatObjectLink: XlClipboardFormatToString = "xlClipboardFormatObjectLink"
        Case xlClipboardFormatBIFF3: XlClipboardFormatToString = "xlClipboardFormatBIFF3"
        Case xlClipboardFormatEmbeddedObject: XlClipboardFormatToString = "xlClipboardFormatEmbeddedObject"
        Case xlClipboardFormatEmbedSource: XlClipboardFormatToString = "xlClipboardFormatEmbedSource"
        Case xlClipboardFormatLinkSource: XlClipboardFormatToString = "xlClipboardFormatLinkSource"
        Case xlClipboardFormatMovie: XlClipboardFormatToString = "xlClipboardFormatMovie"
        Case xlClipboardFormatToolFace: XlClipboardFormatToString = "xlClipboardFormatToolFace"
        Case xlClipboardFormatToolFacePICT: XlClipboardFormatToString = "xlClipboardFormatToolFacePICT"
        Case xlClipboardFormatStandardScale: XlClipboardFormatToString = "xlClipboardFormatStandardScale"
        Case xlClipboardFormatStandardFont: XlClipboardFormatToString = "xlClipboardFormatStandardFont"
        Case xlClipboardFormatScreenPICT: XlClipboardFormatToString = "xlClipboardFormatScreenPICT"
        Case xlClipboardFormatBIFF4: XlClipboardFormatToString = "xlClipboardFormatBIFF4"
        Case xlClipboardFormatObjectDesc: XlClipboardFormatToString = "xlClipboardFormatObjectDesc"
        Case xlClipboardFormatLinkSourceDesc: XlClipboardFormatToString = "xlClipboardFormatLinkSourceDesc"
        Case xlClipboardFormatBIFF12: XlClipboardFormatToString = "xlClipboardFormatBIFF12"
    End Select
End Function
