Attribute VB_Name = "wOlSaveAsType"
Function OlSaveAsTypeFromString(value As String) As OlSaveAsType
    If IsNumeric(value) Then
        OlSaveAsTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olTXT": OlSaveAsTypeFromString = olTXT
        Case "olRTF": OlSaveAsTypeFromString = olRTF
        Case "olTemplate": OlSaveAsTypeFromString = olTemplate
        Case "olMSG": OlSaveAsTypeFromString = olMSG
        Case "olDoc": OlSaveAsTypeFromString = olDoc
        Case "olHTML": OlSaveAsTypeFromString = olHTML
        Case "olVCard": OlSaveAsTypeFromString = olVCard
        Case "olVCal": OlSaveAsTypeFromString = olVCal
        Case "olICal": OlSaveAsTypeFromString = olICal
        Case "olMSGUnicode": OlSaveAsTypeFromString = olMSGUnicode
        Case "olMHTML": OlSaveAsTypeFromString = olMHTML
    End Select
End Function

Function OlSaveAsTypeToString(value As OlSaveAsType) As String
    Select Case value
        Case olTXT: OlSaveAsTypeToString = "olTXT"
        Case olRTF: OlSaveAsTypeToString = "olRTF"
        Case olTemplate: OlSaveAsTypeToString = "olTemplate"
        Case olMSG: OlSaveAsTypeToString = "olMSG"
        Case olDoc: OlSaveAsTypeToString = "olDoc"
        Case olHTML: OlSaveAsTypeToString = "olHTML"
        Case olVCard: OlSaveAsTypeToString = "olVCard"
        Case olVCal: OlSaveAsTypeToString = "olVCal"
        Case olICal: OlSaveAsTypeToString = "olICal"
        Case olMSGUnicode: OlSaveAsTypeToString = "olMSGUnicode"
        Case olMHTML: OlSaveAsTypeToString = "olMHTML"
    End Select
End Function
