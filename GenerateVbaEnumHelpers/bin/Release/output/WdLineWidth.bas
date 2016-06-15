Attribute VB_Name = "wWdLineWidth"
Function WdLineWidthFromString(value As String) As WdLineWidth
    If IsNumeric(value) Then
        WdLineWidthFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdLineWidth025pt": WdLineWidthFromString = wdLineWidth025pt
        Case "wdLineWidth050pt": WdLineWidthFromString = wdLineWidth050pt
        Case "wdLineWidth075pt": WdLineWidthFromString = wdLineWidth075pt
        Case "wdLineWidth100pt": WdLineWidthFromString = wdLineWidth100pt
        Case "wdLineWidth150pt": WdLineWidthFromString = wdLineWidth150pt
        Case "wdLineWidth225pt": WdLineWidthFromString = wdLineWidth225pt
        Case "wdLineWidth300pt": WdLineWidthFromString = wdLineWidth300pt
        Case "wdLineWidth450pt": WdLineWidthFromString = wdLineWidth450pt
        Case "wdLineWidth600pt": WdLineWidthFromString = wdLineWidth600pt
    End Select
End Function

Function WdLineWidthToString(value As WdLineWidth) As String
    Select Case value
        Case wdLineWidth025pt: WdLineWidthToString = "wdLineWidth025pt"
        Case wdLineWidth050pt: WdLineWidthToString = "wdLineWidth050pt"
        Case wdLineWidth075pt: WdLineWidthToString = "wdLineWidth075pt"
        Case wdLineWidth100pt: WdLineWidthToString = "wdLineWidth100pt"
        Case wdLineWidth150pt: WdLineWidthToString = "wdLineWidth150pt"
        Case wdLineWidth225pt: WdLineWidthToString = "wdLineWidth225pt"
        Case wdLineWidth300pt: WdLineWidthToString = "wdLineWidth300pt"
        Case wdLineWidth450pt: WdLineWidthToString = "wdLineWidth450pt"
        Case wdLineWidth600pt: WdLineWidthToString = "wdLineWidth600pt"
    End Select
End Function
