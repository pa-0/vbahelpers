Attribute VB_Name = "wPbSpotColor"
Function PbSpotColorFromString(value As String) As PbSpotColor
    If IsNumeric(value) Then
        PbSpotColorFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "pbInkNone": PbSpotColorFromString = pbInkNone
    End Select
End Function

Function PbSpotColorToString(value As PbSpotColor) As String
    Select Case value
        Case pbInkNone: PbSpotColorToString = "pbInkNone"
    End Select
End Function
