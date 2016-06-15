Attribute VB_Name = "wWdPictureLinkType"
Function WdPictureLinkTypeFromString(value As String) As WdPictureLinkType
    If IsNumeric(value) Then
        WdPictureLinkTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdLinkNone": WdPictureLinkTypeFromString = wdLinkNone
        Case "wdLinkDataInDoc": WdPictureLinkTypeFromString = wdLinkDataInDoc
        Case "wdLinkDataOnDisk": WdPictureLinkTypeFromString = wdLinkDataOnDisk
    End Select
End Function

Function WdPictureLinkTypeToString(value As WdPictureLinkType) As String
    Select Case value
        Case wdLinkNone: WdPictureLinkTypeToString = "wdLinkNone"
        Case wdLinkDataInDoc: WdPictureLinkTypeToString = "wdLinkDataInDoc"
        Case wdLinkDataOnDisk: WdPictureLinkTypeToString = "wdLinkDataOnDisk"
    End Select
End Function
