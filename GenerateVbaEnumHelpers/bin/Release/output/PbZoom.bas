Attribute VB_Name = "wPbZoom"
Function PbZoomFromString(value As String) As PbZoom
    If IsNumeric(value) Then
        PbZoomFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbZoomFitSelection": PbZoomFromString = pbZoomFitSelection
        Case "pbZoomWholePage": PbZoomFromString = pbZoomWholePage
        Case "pbZoomPageWidth": PbZoomFromString = pbZoomPageWidth
    End Select
End Function

Function PbZoomToString(value As PbZoom) As String
    Select Case value
        Case pbZoomFitSelection: PbZoomToString = "pbZoomFitSelection"
        Case pbZoomWholePage: PbZoomToString = "pbZoomWholePage"
        Case pbZoomPageWidth: PbZoomToString = "pbZoomPageWidth"
    End Select
End Function
