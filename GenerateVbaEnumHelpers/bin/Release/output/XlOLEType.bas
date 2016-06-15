Attribute VB_Name = "wXlOLEType"
Function XlOLETypeFromString(value As String) As XlOLEType
    If IsNumeric(value) Then
        XlOLETypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlOLELink": XlOLETypeFromString = xlOLELink
        Case "xlOLEEmbed": XlOLETypeFromString = xlOLEEmbed
        Case "xlOLEControl": XlOLETypeFromString = xlOLEControl
    End Select
End Function

Function XlOLETypeToString(value As XlOLEType) As String
    Select Case value
        Case xlOLELink: XlOLETypeToString = "xlOLELink"
        Case xlOLEEmbed: XlOLETypeToString = "xlOLEEmbed"
        Case xlOLEControl: XlOLETypeToString = "xlOLEControl"
    End Select
End Function
