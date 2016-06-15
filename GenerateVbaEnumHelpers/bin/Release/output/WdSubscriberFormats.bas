Attribute VB_Name = "wWdSubscriberFormats"
Function WdSubscriberFormatsFromString(value As String) As WdSubscriberFormats
    If IsNumeric(value) Then
        WdSubscriberFormatsFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdSubscriberBestFormat": WdSubscriberFormatsFromString = wdSubscriberBestFormat
        Case "wdSubscriberRTF": WdSubscriberFormatsFromString = wdSubscriberRTF
        Case "wdSubscriberText": WdSubscriberFormatsFromString = wdSubscriberText
        Case "wdSubscriberPict": WdSubscriberFormatsFromString = wdSubscriberPict
    End Select
End Function

Function WdSubscriberFormatsToString(value As WdSubscriberFormats) As String
    Select Case value
        Case wdSubscriberBestFormat: WdSubscriberFormatsToString = "wdSubscriberBestFormat"
        Case wdSubscriberRTF: WdSubscriberFormatsToString = "wdSubscriberRTF"
        Case wdSubscriberText: WdSubscriberFormatsToString = "wdSubscriberText"
        Case wdSubscriberPict: WdSubscriberFormatsToString = "wdSubscriberPict"
    End Select
End Function
