Attribute VB_Name = "wMsoAutoSize"
Function MsoAutoSizeFromString(value As String) As MsoAutoSize
    If IsNumeric(value) Then
        MsoAutoSizeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "msoAutoSizeNone": MsoAutoSizeFromString = msoAutoSizeNone
        Case "msoAutoSizeShapeToFitText": MsoAutoSizeFromString = msoAutoSizeShapeToFitText
        Case "msoAutoSizeTextToFitShape": MsoAutoSizeFromString = msoAutoSizeTextToFitShape
        Case "msoAutoSizeMixed": MsoAutoSizeFromString = msoAutoSizeMixed
    End Select
End Function

Function MsoAutoSizeToString(value As MsoAutoSize) As String
    Select Case value
        Case msoAutoSizeNone: MsoAutoSizeToString = "msoAutoSizeNone"
        Case msoAutoSizeShapeToFitText: MsoAutoSizeToString = "msoAutoSizeShapeToFitText"
        Case msoAutoSizeTextToFitShape: MsoAutoSizeToString = "msoAutoSizeTextToFitShape"
        Case msoAutoSizeMixed: MsoAutoSizeToString = "msoAutoSizeMixed"
    End Select
End Function
