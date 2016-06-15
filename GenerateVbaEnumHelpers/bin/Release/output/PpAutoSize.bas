Attribute VB_Name = "wPpAutoSize"
Function PpAutoSizeFromString(value As String) As PpAutoSize
    If IsNumeric(value) Then
        PpAutoSizeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "ppAutoSizeNone": PpAutoSizeFromString = ppAutoSizeNone
        Case "ppAutoSizeShapeToFitText": PpAutoSizeFromString = ppAutoSizeShapeToFitText
        Case "ppAutoSizeMixed": PpAutoSizeFromString = ppAutoSizeMixed
    End Select
End Function

Function PpAutoSizeToString(value As PpAutoSize) As String
    Select Case value
        Case ppAutoSizeNone: PpAutoSizeToString = "ppAutoSizeNone"
        Case ppAutoSizeShapeToFitText: PpAutoSizeToString = "ppAutoSizeShapeToFitText"
        Case ppAutoSizeMixed: PpAutoSizeToString = "ppAutoSizeMixed"
    End Select
End Function
