Attribute VB_Name = "wWdCaptionLabelID"
Function WdCaptionLabelIDFromString(value As String) As WdCaptionLabelID
    If IsNumeric(value) Then
        WdCaptionLabelIDFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "wdCaptionEquation": WdCaptionLabelIDFromString = wdCaptionEquation
        Case "wdCaptionTable": WdCaptionLabelIDFromString = wdCaptionTable
        Case "wdCaptionFigure": WdCaptionLabelIDFromString = wdCaptionFigure
    End Select
End Function

Function WdCaptionLabelIDToString(value As WdCaptionLabelID) As String
    Select Case value
        Case wdCaptionEquation: WdCaptionLabelIDToString = "wdCaptionEquation"
        Case wdCaptionTable: WdCaptionLabelIDToString = "wdCaptionTable"
        Case wdCaptionFigure: WdCaptionLabelIDToString = "wdCaptionFigure"
    End Select
End Function
