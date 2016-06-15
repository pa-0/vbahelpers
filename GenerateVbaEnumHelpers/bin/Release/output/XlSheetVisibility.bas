Attribute VB_Name = "wXlSheetVisibility"
Function XlSheetVisibilityFromString(value As String) As XlSheetVisibility
    If IsNumeric(value) Then
        XlSheetVisibilityFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlSheetHidden": XlSheetVisibilityFromString = xlSheetHidden
        Case "xlSheetVeryHidden": XlSheetVisibilityFromString = xlSheetVeryHidden
        Case "xlSheetVisible": XlSheetVisibilityFromString = xlSheetVisible
    End Select
End Function

Function XlSheetVisibilityToString(value As XlSheetVisibility) As String
    Select Case value
        Case xlSheetHidden: XlSheetVisibilityToString = "xlSheetHidden"
        Case xlSheetVeryHidden: XlSheetVisibilityToString = "xlSheetVeryHidden"
        Case xlSheetVisible: XlSheetVisibilityToString = "xlSheetVisible"
    End Select
End Function
