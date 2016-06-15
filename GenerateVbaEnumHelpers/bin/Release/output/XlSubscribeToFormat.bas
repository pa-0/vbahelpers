Attribute VB_Name = "wXlSubscribeToFormat"
Function XlSubscribeToFormatFromString(value As String) As XlSubscribeToFormat
    If IsNumeric(value) Then
        XlSubscribeToFormatFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlSubscribeToText": XlSubscribeToFormatFromString = xlSubscribeToText
        Case "xlSubscribeToPicture": XlSubscribeToFormatFromString = xlSubscribeToPicture
    End Select
End Function

Function XlSubscribeToFormatToString(value As XlSubscribeToFormat) As String
    Select Case value
        Case xlSubscribeToText: XlSubscribeToFormatToString = "xlSubscribeToText"
        Case xlSubscribeToPicture: XlSubscribeToFormatToString = "xlSubscribeToPicture"
    End Select
End Function
