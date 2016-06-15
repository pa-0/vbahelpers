Attribute VB_Name = "wOlEditorType"
Function OlEditorTypeFromString(value As String) As OlEditorType
    If IsNumeric(value) Then
        OlEditorTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olEditorText": OlEditorTypeFromString = olEditorText
        Case "olEditorHTML": OlEditorTypeFromString = olEditorHTML
        Case "olEditorRTF": OlEditorTypeFromString = olEditorRTF
        Case "olEditorWord": OlEditorTypeFromString = olEditorWord
    End Select
End Function

Function OlEditorTypeToString(value As OlEditorType) As String
    Select Case value
        Case olEditorText: OlEditorTypeToString = "olEditorText"
        Case olEditorHTML: OlEditorTypeToString = "olEditorHTML"
        Case olEditorRTF: OlEditorTypeToString = "olEditorRTF"
        Case olEditorWord: OlEditorTypeToString = "olEditorWord"
    End Select
End Function
