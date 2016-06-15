Attribute VB_Name = "wXlDisplayBlanksAs"
Function XlDisplayBlanksAsFromString(value As String) As XlDisplayBlanksAs
    If IsNumeric(value) Then
        XlDisplayBlanksAsFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "xlNotPlotted": XlDisplayBlanksAsFromString = xlNotPlotted
        Case "xlZero": XlDisplayBlanksAsFromString = xlZero
        Case "xlInterpolated": XlDisplayBlanksAsFromString = xlInterpolated
    End Select
End Function

Function XlDisplayBlanksAsToString(value As XlDisplayBlanksAs) As String
    Select Case value
        Case xlNotPlotted: XlDisplayBlanksAsToString = "xlNotPlotted"
        Case xlZero: XlDisplayBlanksAsToString = "xlZero"
        Case xlInterpolated: XlDisplayBlanksAsToString = "xlInterpolated"
    End Select
End Function
