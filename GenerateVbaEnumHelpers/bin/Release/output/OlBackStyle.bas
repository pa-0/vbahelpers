Attribute VB_Name = "wOlBackStyle"
Function OlBackStyleFromString(value As String) As OlBackStyle
    If IsNumeric(value) Then
        OlBackStyleFromString = CInt(value)
        Exit Function
    End If

    Select Case value
        Case "olBackStyleTransparent": OlBackStyleFromString = olBackStyleTransparent
        Case "olBackStyleOpaque": OlBackStyleFromString = olBackStyleOpaque
    End Select
End Function

Function OlBackStyleToString(value As OlBackStyle) As String
    Select Case value
        Case olBackStyleTransparent: OlBackStyleToString = "olBackStyleTransparent"
        Case olBackStyleOpaque: OlBackStyleToString = "olBackStyleOpaque"
    End Select
End Function
