Attribute VB_Name = "wMsoHyperlinkType"
Function MsoHyperlinkTypeFromString(value As String) As MsoHyperlinkType
    If IsNumeric(value) Then
        MsoHyperlinkTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoHyperlinkRange": MsoHyperlinkTypeFromString = msoHyperlinkRange
        Case "msoHyperlinkShape": MsoHyperlinkTypeFromString = msoHyperlinkShape
        Case "msoHyperlinkInlineShape": MsoHyperlinkTypeFromString = msoHyperlinkInlineShape
    End Select
End Function

Function MsoHyperlinkTypeToString(value As MsoHyperlinkType) As String
    Select Case value
        Case msoHyperlinkRange: MsoHyperlinkTypeToString = "msoHyperlinkRange"
        Case msoHyperlinkShape: MsoHyperlinkTypeToString = "msoHyperlinkShape"
        Case msoHyperlinkInlineShape: MsoHyperlinkTypeToString = "msoHyperlinkInlineShape"
    End Select
End Function
