Attribute VB_Name = "wXlFindLookIn"
Function XlFindLookInFromString(value As String) As XlFindLookIn
    If IsNumeric(value) Then
        XlFindLookInFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlValues": XlFindLookInFromString = xlValues
        Case "xlComments": XlFindLookInFromString = xlComments
        Case "xlFormulas": XlFindLookInFromString = xlFormulas
    End Select
End Function

Function XlFindLookInToString(value As XlFindLookIn) As String
    Select Case value
        Case xlValues: XlFindLookInToString = "xlValues"
        Case xlComments: XlFindLookInToString = "xlComments"
        Case xlFormulas: XlFindLookInToString = "xlFormulas"
    End Select
End Function
