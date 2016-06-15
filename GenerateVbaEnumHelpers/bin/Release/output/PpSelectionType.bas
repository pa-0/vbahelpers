Attribute VB_Name = "wPpSelectionType"
Function PpSelectionTypeFromString(value As String) As PpSelectionType
    If IsNumeric(value) Then
        PpSelectionTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "ppSelectionNone": PpSelectionTypeFromString = ppSelectionNone
        Case "ppSelectionSlides": PpSelectionTypeFromString = ppSelectionSlides
        Case "ppSelectionShapes": PpSelectionTypeFromString = ppSelectionShapes
        Case "ppSelectionText": PpSelectionTypeFromString = ppSelectionText
    End Select
End Function

Function PpSelectionTypeToString(value As PpSelectionType) As String
    Select Case value
        Case ppSelectionNone: PpSelectionTypeToString = "ppSelectionNone"
        Case ppSelectionSlides: PpSelectionTypeToString = "ppSelectionSlides"
        Case ppSelectionShapes: PpSelectionTypeToString = "ppSelectionShapes"
        Case ppSelectionText: PpSelectionTypeToString = "ppSelectionText"
    End Select
End Function
