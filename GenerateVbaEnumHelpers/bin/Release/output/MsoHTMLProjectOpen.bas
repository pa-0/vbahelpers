Attribute VB_Name = "wMsoHTMLProjectOpen"
Function MsoHTMLProjectOpenFromString(value As String) As MsoHTMLProjectOpen
    If IsNumeric(value) Then
        MsoHTMLProjectOpenFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoHTMLProjectOpenSourceView": MsoHTMLProjectOpenFromString = msoHTMLProjectOpenSourceView
        Case "msoHTMLProjectOpenTextView": MsoHTMLProjectOpenFromString = msoHTMLProjectOpenTextView
    End Select
End Function

Function MsoHTMLProjectOpenToString(value As MsoHTMLProjectOpen) As String
    Select Case value
        Case msoHTMLProjectOpenSourceView: MsoHTMLProjectOpenToString = "msoHTMLProjectOpenSourceView"
        Case msoHTMLProjectOpenTextView: MsoHTMLProjectOpenToString = "msoHTMLProjectOpenTextView"
    End Select
End Function
