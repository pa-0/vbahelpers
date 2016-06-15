Attribute VB_Name = "wMsoCommandBarButtonHyperlinkType"
Function MsoCommandBarButtonHyperlinkTypeFromString(value As String) As MsoCommandBarButtonHyperlinkType
    If IsNumeric(value) Then
        MsoCommandBarButtonHyperlinkTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoCommandBarButtonHyperlinkNone": MsoCommandBarButtonHyperlinkTypeFromString = msoCommandBarButtonHyperlinkNone
        Case "msoCommandBarButtonHyperlinkOpen": MsoCommandBarButtonHyperlinkTypeFromString = msoCommandBarButtonHyperlinkOpen
        Case "msoCommandBarButtonHyperlinkInsertPicture": MsoCommandBarButtonHyperlinkTypeFromString = msoCommandBarButtonHyperlinkInsertPicture
    End Select
End Function

Function MsoCommandBarButtonHyperlinkTypeToString(value As MsoCommandBarButtonHyperlinkType) As String
    Select Case value
        Case msoCommandBarButtonHyperlinkNone: MsoCommandBarButtonHyperlinkTypeToString = "msoCommandBarButtonHyperlinkNone"
        Case msoCommandBarButtonHyperlinkOpen: MsoCommandBarButtonHyperlinkTypeToString = "msoCommandBarButtonHyperlinkOpen"
        Case msoCommandBarButtonHyperlinkInsertPicture: MsoCommandBarButtonHyperlinkTypeToString = "msoCommandBarButtonHyperlinkInsertPicture"
    End Select
End Function
