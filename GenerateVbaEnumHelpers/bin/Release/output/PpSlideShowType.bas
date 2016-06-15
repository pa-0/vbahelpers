Attribute VB_Name = "wPpSlideShowType"
Function PpSlideShowTypeFromString(value As String) As PpSlideShowType
    If IsNumeric(value) Then
        PpSlideShowTypeFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "ppShowTypeSpeaker": PpSlideShowTypeFromString = ppShowTypeSpeaker
        Case "ppShowTypeWindow": PpSlideShowTypeFromString = ppShowTypeWindow
        Case "ppShowTypeKiosk": PpSlideShowTypeFromString = ppShowTypeKiosk
        Case "ppShowTypeWindow2": PpSlideShowTypeFromString = ppShowTypeWindow2
    End Select
End Function

Function PpSlideShowTypeToString(value As PpSlideShowType) As String
    Select Case value
        Case ppShowTypeSpeaker: PpSlideShowTypeToString = "ppShowTypeSpeaker"
        Case ppShowTypeWindow: PpSlideShowTypeToString = "ppShowTypeWindow"
        Case ppShowTypeKiosk: PpSlideShowTypeToString = "ppShowTypeKiosk"
        Case ppShowTypeWindow2: PpSlideShowTypeToString = "ppShowTypeWindow2"
    End Select
End Function
