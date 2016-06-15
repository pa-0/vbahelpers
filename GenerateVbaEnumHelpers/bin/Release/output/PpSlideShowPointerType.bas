Attribute VB_Name = "wPpSlideShowPointerType"
Function PpSlideShowPointerTypeFromString(value As String) As PpSlideShowPointerType
    If IsNumeric(value) Then
        PpSlideShowPointerTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "ppSlideShowPointerNone": PpSlideShowPointerTypeFromString = ppSlideShowPointerNone
        Case "ppSlideShowPointerArrow": PpSlideShowPointerTypeFromString = ppSlideShowPointerArrow
        Case "ppSlideShowPointerPen": PpSlideShowPointerTypeFromString = ppSlideShowPointerPen
        Case "ppSlideShowPointerAlwaysHidden": PpSlideShowPointerTypeFromString = ppSlideShowPointerAlwaysHidden
        Case "ppSlideShowPointerAutoArrow": PpSlideShowPointerTypeFromString = ppSlideShowPointerAutoArrow
        Case "ppSlideShowPointerEraser": PpSlideShowPointerTypeFromString = ppSlideShowPointerEraser
    End Select
End Function

Function PpSlideShowPointerTypeToString(value As PpSlideShowPointerType) As String
    Select Case value
        Case ppSlideShowPointerNone: PpSlideShowPointerTypeToString = "ppSlideShowPointerNone"
        Case ppSlideShowPointerArrow: PpSlideShowPointerTypeToString = "ppSlideShowPointerArrow"
        Case ppSlideShowPointerPen: PpSlideShowPointerTypeToString = "ppSlideShowPointerPen"
        Case ppSlideShowPointerAlwaysHidden: PpSlideShowPointerTypeToString = "ppSlideShowPointerAlwaysHidden"
        Case ppSlideShowPointerAutoArrow: PpSlideShowPointerTypeToString = "ppSlideShowPointerAutoArrow"
        Case ppSlideShowPointerEraser: PpSlideShowPointerTypeToString = "ppSlideShowPointerEraser"
    End Select
End Function
