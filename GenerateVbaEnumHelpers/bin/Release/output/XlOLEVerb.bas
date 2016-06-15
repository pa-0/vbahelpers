Attribute VB_Name = "wXlOLEVerb"
Function XlOLEVerbFromString(value As String) As XlOLEVerb
    If IsNumeric(value) Then
        XlOLEVerbFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlVerbPrimary": XlOLEVerbFromString = xlVerbPrimary
        Case "xlVerbOpen": XlOLEVerbFromString = xlVerbOpen
    End Select
End Function

Function XlOLEVerbToString(value As XlOLEVerb) As String
    Select Case value
        Case xlVerbPrimary: XlOLEVerbToString = "xlVerbPrimary"
        Case xlVerbOpen: XlOLEVerbToString = "xlVerbOpen"
    End Select
End Function
