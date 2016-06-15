Attribute VB_Name = "wOlAttachmentType"
Function OlAttachmentTypeFromString(value As String) As OlAttachmentType
    If IsNumeric(value) Then
        OlAttachmentTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olByValue": OlAttachmentTypeFromString = olByValue
        Case "olByReference": OlAttachmentTypeFromString = olByReference
        Case "olEmbeddeditem": OlAttachmentTypeFromString = olEmbeddeditem
        Case "olOLE": OlAttachmentTypeFromString = olOLE
    End Select
End Function

Function OlAttachmentTypeToString(value As OlAttachmentType) As String
    Select Case value
        Case olByValue: OlAttachmentTypeToString = "olByValue"
        Case olByReference: OlAttachmentTypeToString = "olByReference"
        Case olEmbeddeditem: OlAttachmentTypeToString = "olEmbeddeditem"
        Case olOLE: OlAttachmentTypeToString = "olOLE"
    End Select
End Function
