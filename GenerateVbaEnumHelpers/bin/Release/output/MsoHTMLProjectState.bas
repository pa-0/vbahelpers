Attribute VB_Name = "wMsoHTMLProjectState"
Function MsoHTMLProjectStateFromString(value As String) As MsoHTMLProjectState
    If IsNumeric(value) Then
        MsoHTMLProjectStateFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoHTMLProjectStateDocumentLocked": MsoHTMLProjectStateFromString = msoHTMLProjectStateDocumentLocked
        Case "msoHTMLProjectStateProjectLocked": MsoHTMLProjectStateFromString = msoHTMLProjectStateProjectLocked
        Case "msoHTMLProjectStateDocumentProjectUnlocked": MsoHTMLProjectStateFromString = msoHTMLProjectStateDocumentProjectUnlocked
    End Select
End Function

Function MsoHTMLProjectStateToString(value As MsoHTMLProjectState) As String
    Select Case value
        Case msoHTMLProjectStateDocumentLocked: MsoHTMLProjectStateToString = "msoHTMLProjectStateDocumentLocked"
        Case msoHTMLProjectStateProjectLocked: MsoHTMLProjectStateToString = "msoHTMLProjectStateProjectLocked"
        Case msoHTMLProjectStateDocumentProjectUnlocked: MsoHTMLProjectStateToString = "msoHTMLProjectStateDocumentProjectUnlocked"
    End Select
End Function
